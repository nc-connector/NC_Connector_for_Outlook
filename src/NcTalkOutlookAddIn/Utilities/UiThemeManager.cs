/*
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using Microsoft.Win32;
using NcTalkOutlookAddIn.UI;

namespace NcTalkOutlookAddIn.Utilities
{
    internal enum UiThemeKind
    {
        Light,
        Dark
    }

    internal enum OfficeUiTheme
    {
        Unknown = -1,
        Colorful = 0,
        DarkGray = 1,
        Black = 2,
        White = 3
    }

    internal sealed class UiThemePalette
    {
        internal UiThemePalette(UiThemeKind kind, OfficeUiTheme officeUiTheme = OfficeUiTheme.Unknown)
        {
            Kind = kind;
            OfficeUiTheme = officeUiTheme;
            IsDark = kind == UiThemeKind.Dark;

            if (IsDark)
            {
                // Outlook/Office dark themes vary ("Dark Gray" vs "Black"). We aim for a close match
                // while keeping readability across Windows builds. Defaults lean towards "Black".
                bool useOfficeBlack = officeUiTheme == OfficeUiTheme.Black || officeUiTheme == OfficeUiTheme.Unknown;
                bool useOfficeDarkGray = officeUiTheme == OfficeUiTheme.DarkGray;

                if (useOfficeDarkGray)
                {
                    WindowBackground = Color.FromArgb(58, 58, 58);
                    ControlBackground = Color.FromArgb(68, 68, 68);
                    InputBackground = Color.FromArgb(80, 80, 80);
                    Border = Color.FromArgb(98, 98, 98);
                    MutedText = Color.FromArgb(200, 200, 200);
                    SelectionBackground = Color.FromArgb(88, 88, 88);
                }
                else if (useOfficeBlack)
                {
                    WindowBackground = Color.FromArgb(32, 32, 32);
                    ControlBackground = Color.FromArgb(43, 43, 43);
                    InputBackground = Color.FromArgb(55, 55, 55);
                    Border = Color.FromArgb(74, 74, 74);
                    MutedText = Color.FromArgb(190, 190, 190);
                    SelectionBackground = Color.FromArgb(64, 64, 64);
                }
                else
                {
                    WindowBackground = Color.FromArgb(30, 30, 30);
                    ControlBackground = Color.FromArgb(37, 37, 38);
                    InputBackground = Color.FromArgb(45, 45, 48);
                    Border = Color.FromArgb(70, 70, 70);
                    MutedText = Color.FromArgb(170, 170, 170);
                    SelectionBackground = Color.FromArgb(62, 62, 66);
                }

                Text = Color.FromArgb(242, 242, 242);
                LinkText = Color.FromArgb(0, 130, 201); // brand blue
                SuccessText = Color.FromArgb(110, 200, 110);
                ErrorText = Color.FromArgb(230, 120, 120);
                WarningText = Color.FromArgb(230, 200, 120);
                SelectionText = Text;
                DisabledText = Color.FromArgb(160, 160, 160);
                AvatarPlaceholderFill = Color.FromArgb(70, 70, 70);
                AvatarPlaceholderBorder = Border;
                AvatarPlaceholderText = Text;
            }
            else
            {
                WindowBackground = SystemColors.Window;
                ControlBackground = SystemColors.Control;
                InputBackground = SystemColors.Window;
                Border = SystemColors.ControlDark;
                Text = SystemColors.ControlText;
                MutedText = Color.DimGray;
                LinkText = Color.FromArgb(0, 130, 201);
                SuccessText = Color.DarkGreen;
                ErrorText = Color.DarkRed;
                WarningText = Color.DarkGoldenrod;
                SelectionBackground = SystemColors.Highlight;
                SelectionText = SystemColors.HighlightText;
                DisabledText = SystemColors.GrayText;
                AvatarPlaceholderFill = Color.FromArgb(220, 220, 220);
                AvatarPlaceholderBorder = Color.FromArgb(160, 160, 160);
                AvatarPlaceholderText = Color.FromArgb(70, 70, 70);
            }
        }

        internal UiThemeKind Kind { get; private set; }

        internal OfficeUiTheme OfficeUiTheme { get; private set; }

        internal bool IsDark { get; private set; }

        internal Color WindowBackground { get; private set; }

        internal Color ControlBackground { get; private set; }

        internal Color InputBackground { get; private set; }

        internal Color Border { get; private set; }

        internal Color Text { get; private set; }

        internal Color MutedText { get; private set; }

        internal Color LinkText { get; private set; }

        internal Color SuccessText { get; private set; }

        internal Color ErrorText { get; private set; }

        internal Color WarningText { get; private set; }

        internal Color SelectionBackground { get; private set; }

        internal Color SelectionText { get; private set; }

        internal Color DisabledText { get; private set; }

        internal Color AvatarPlaceholderFill { get; private set; }

        internal Color AvatarPlaceholderBorder { get; private set; }

        internal Color AvatarPlaceholderText { get; private set; }
    }

    internal static class UiThemeManager
    {
        private static readonly ConditionalWeakTable<TabControl, object> TabStretchApplied =
            new ConditionalWeakTable<TabControl, object>();
        private static readonly ConditionalWeakTable<TabControl, object> TabChromeApplied =
            new ConditionalWeakTable<TabControl, object>();

        internal static UiThemePalette DetectPalette()
        {
            if (SystemInformation.HighContrast)
            {
                return new UiThemePalette(UiThemeKind.Light);
            }

            OfficeUiTheme? officeTheme = OutlookThemeDetector.TryGetOfficeUiTheme();
            if (officeTheme.HasValue)
            {
                if (officeTheme.Value == OfficeUiTheme.Black || officeTheme.Value == OfficeUiTheme.DarkGray)
                {
                    return new UiThemePalette(UiThemeKind.Dark, officeTheme.Value);
                }

                return new UiThemePalette(UiThemeKind.Light, officeTheme.Value);
            }

            return OutlookThemeDetector.IsDarkThemePreferred()
                ? new UiThemePalette(UiThemeKind.Dark, OfficeUiTheme.Unknown)
                : new UiThemePalette(UiThemeKind.Light, OfficeUiTheme.Unknown);
        }

        internal static void ApplyToForm(Form form, params ToolTip[] toolTips)
        {
            if (form == null)
            {
                return;
            }

            UiThemePalette palette = DetectPalette();

            // Tab "stretching" is a layout preference, not a theme. Apply it regardless of light/dark
            // so the settings tabs always fill the available width (no leftover strip behind the last tab).
            try
            {
                ApplyTabStretchToControlTree(form);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to apply tab stretching to control tree.", ex);
            }

            // The settings tab strip requires a custom chrome to avoid bright theme bleed-through.
            // Apply it for both light and dark palettes.
            try
            {
                ApplyTabChromeToControlTree(form, palette);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to apply tab chrome to control tree.", ex);
            }

            if (!palette.IsDark)
            {
                return;
            }

            ApplyToControlTree(form, palette);

            if (toolTips != null)
            {
                foreach (var toolTip in toolTips)
                {
                    ApplyToToolTip(toolTip, palette);
                }
            }

            if (form.IsHandleCreated)
            {
                TrySetDarkTitleBar(form.Handle, true);
            }
            else
            {
                form.HandleCreated += (s, e) => TrySetDarkTitleBar(form.Handle, true);
            }

            // Some common controls (TabControl, ComboBox drop-down, etc.) may reset colors after their handles are created.
            // Re-apply once the form is shown to ensure a consistent palette.
            form.Shown += (s, e) =>
            {
                try
                {
                    ApplyToControlTree(form, palette);
                    form.Invalidate(true);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to re-apply theme on form shown.", ex);
                }
            };
        }

        internal static void ApplyToControlTree(Control root, UiThemePalette palette)
        {
            if (root == null || palette == null || !palette.IsDark)
            {
                return;
            }

            ApplyControlColors(root, palette);

            if (root is ListView)
            {
                ApplyListView((ListView)root, palette);
            }

            if (root is DateTimePicker)
            {
                ApplyDateTimePicker((DateTimePicker)root, palette);
            }

            if (root is ComboBox)
            {
                ApplyComboBox((ComboBox)root, palette);
            }

            foreach (Control child in root.Controls)
            {
                if (child is BrandedHeader)
                {
                    continue;
                }

                ApplyToControlTree(child, palette);
            }
        }

        private static void ApplyTabChromeToControlTree(Control root, UiThemePalette palette)
        {
            if (root == null || palette == null)
            {
                return;
            }

            if (root is TabControl)
            {
                ApplyTabControl((TabControl)root, palette);
            }

            foreach (Control child in root.Controls)
            {
                if (child is BrandedHeader)
                {
                    continue;
                }

                ApplyTabChromeToControlTree(child, palette);
            }
        }

        private static void ApplyControlColors(Control control, UiThemePalette palette)
        {
            if (control == null || palette == null || !palette.IsDark)
            {
                return;
            }

            if (control is BrandedHeader)
            {
                return;
            }

            if (control is Form)
            {
                control.BackColor = palette.WindowBackground;
                control.ForeColor = palette.Text;
                return;
            }

            if (control is TabPage || control is Panel)
            {
                control.BackColor = palette.WindowBackground;
                control.ForeColor = palette.Text;
                return;
            }

            if (control is GroupBox)
            {
                control.BackColor = palette.WindowBackground;
                // The standard GroupBox uses ForeColor for both the caption and the border. A slightly muted tone
                // looks closer to Outlook's dark UI and avoids overly bright borders.
                control.ForeColor = palette.MutedText;
                return;
            }

            if (control is Label)
            {
                var label = (Label)control;
                if (label.ForeColor == Color.DimGray || label.ForeColor == Color.FromArgb(64, 64, 64))
                {
                    label.ForeColor = palette.MutedText;
                }
                else if (label.ForeColor == SystemColors.ControlText || label.ForeColor == Color.Black)
                {
                    label.ForeColor = palette.Text;
                }
                label.BackColor = palette.WindowBackground;
                return;
            }

            if (control is LinkLabel)
            {
                var link = (LinkLabel)control;
                link.BackColor = palette.WindowBackground;
                link.ForeColor = palette.Text;
                link.LinkColor = palette.LinkText;
                link.ActiveLinkColor = palette.LinkText;
                link.VisitedLinkColor = palette.LinkText;
                return;
            }

            if (control is TextBoxBase || control is ComboBox || control is ListBox || control is NumericUpDown)
            {
                control.BackColor = palette.InputBackground;
                control.ForeColor = control.Enabled ? palette.Text : palette.DisabledText;
                return;
            }

            if (control is Button)
            {
                var button = (Button)control;
                button.UseVisualStyleBackColor = false;
                button.BackColor = palette.ControlBackground;
                button.ForeColor = palette.Text;
                button.FlatStyle = FlatStyle.Flat;
                button.FlatAppearance.BorderColor = palette.Border;
                button.FlatAppearance.BorderSize = 1;
                button.FlatAppearance.MouseOverBackColor = palette.SelectionBackground;
                button.FlatAppearance.MouseDownBackColor = palette.SelectionBackground;
                return;
            }

            if (control is CheckBox || control is RadioButton)
            {
                control.BackColor = palette.WindowBackground;
                control.ForeColor = palette.Text;
                return;
            }

            if (control is ListView)
            {
                control.BackColor = palette.InputBackground;
                control.ForeColor = palette.Text;
                return;
            }
        }

        private static void ApplyToToolTip(ToolTip toolTip, UiThemePalette palette)
        {
            if (toolTip == null || palette == null || !palette.IsDark)
            {
                return;
            }

            try
            {
                toolTip.BackColor = palette.ControlBackground;
                toolTip.ForeColor = palette.Text;
            }
            catch (Exception ex)
            {
                // ToolTip theming is optional; never abort UI initialization.
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to apply ToolTip theme.", ex);
            }
        }

        private static void ApplyDateTimePicker(DateTimePicker picker, UiThemePalette palette)
        {
            if (picker == null || palette == null || !palette.IsDark)
            {
                return;
            }

            picker.ForeColor = palette.Text;
            picker.CalendarForeColor = palette.Text;
            picker.CalendarMonthBackground = palette.WindowBackground;
            picker.CalendarTitleBackColor = palette.ControlBackground;
            picker.CalendarTitleForeColor = palette.Text;
            picker.CalendarTrailingForeColor = palette.MutedText;
        }

        private static void ApplyListView(ListView view, UiThemePalette palette)
        {
            if (view == null || palette == null || !palette.IsDark)
            {
                return;
            }

            view.BackColor = palette.InputBackground;
            view.ForeColor = palette.Text;
        }

        private static void ApplyComboBox(ComboBox comboBox, UiThemePalette palette)
        {
            if (comboBox == null || palette == null || !palette.IsDark)
            {
                return;
            }

            comboBox.BackColor = palette.InputBackground;
            comboBox.ForeColor = palette.Text;

            if (comboBox.DrawMode != DrawMode.OwnerDrawFixed)
            {
                comboBox.DrawMode = DrawMode.OwnerDrawFixed;
                comboBox.DrawItem += (sender, e) =>
                {
                    try
                    {
                        var combo = (ComboBox)sender;
                        e.DrawBackground();

                        if (e.Index < 0 || e.Index >= combo.Items.Count)
                        {
                            return;
                        }

                        bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
                        Color back = selected ? palette.SelectionBackground : palette.InputBackground;
                        Color text = selected ? palette.SelectionText : palette.Text;

                        using (var brush = new SolidBrush(back))
                        {
                            e.Graphics.FillRectangle(brush, e.Bounds);
                        }

                        string itemText = combo.GetItemText(combo.Items[e.Index]);
                        TextRenderer.DrawText(
                            e.Graphics,
                            itemText,
                            combo.Font,
                            e.Bounds,
                            text,
                            TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.EndEllipsis);

                        e.DrawFocusRectangle();
                    }
                    catch (Exception ex)
                    {
                        DiagnosticsLogger.LogException(LogCategories.Core, "Failed to owner-draw ComboBox item.", ex);
                    }
                };
            }
        }

        private static void ApplyTabControl(TabControl control, UiThemePalette palette)
        {
            if (control == null || palette == null)
            {
                return;
            }

            ApplyTabStretch(control);

            control.Padding = new Point(0, 0);
            control.Appearance = TabAppearance.Normal;
            control.Multiline = false;

            control.BackColor = palette.WindowBackground;
            control.ForeColor = palette.Text;

            foreach (TabPage page in control.TabPages)
            {
                page.BackColor = palette.WindowBackground;
                page.ForeColor = palette.Text;
                page.UseVisualStyleBackColor = false;
            }

            if (control.IsHandleCreated)
            {
                DisableWindowsTheme(control);
            }
            else
            {
                control.HandleCreated += (s, e) => DisableWindowsTheme(control);
            }

            object existing;
            if (TabChromeApplied.TryGetValue(control, out existing))
            {
                return;
            }

            TabChromeApplied.Add(control, new object());

            if (control.DrawMode != TabDrawMode.OwnerDrawFixed)
            {
                control.DrawMode = TabDrawMode.OwnerDrawFixed;
            }

            control.DrawItem += (sender, e) =>
            {
                try
                {
                    var tab = (TabControl)sender;
                    bool selected = e.Index == tab.SelectedIndex;
                    bool isLast = e.Index == tab.TabPages.Count - 1;
                    Rectangle bounds = e.Bounds;

                     // Slight overlap to cover 1px gaps between tab rectangles.
                     if (!isLast)
                     {
                         bounds.Width = Math.Min(bounds.Width + 1, Math.Max(0, tab.ClientSize.Width - bounds.Left));
                     }
                    else
                    {
                        // Visually extend the last tab to the right edge so there is no "free strip" after the
                        // last label (WinForms keeps a small inset on the right).
                        bounds.Width = Math.Max(bounds.Width, Math.Max(0, tab.ClientSize.Width - bounds.Left));
                    }

                    Color selectedBack = palette.IsDark ? palette.ControlBackground : palette.WindowBackground;
                    Color unselectedBack = palette.IsDark ? palette.WindowBackground : palette.ControlBackground;
                    Color back = selected ? selectedBack : unselectedBack;
                    Color text = selected ? palette.Text : palette.MutedText;

                    using (var brush = new SolidBrush(back))
                    {
                        e.Graphics.FillRectangle(brush, bounds);
                    }

                    using (var borderPen = new Pen(palette.Border))
                    {
                        // Draw only separators between tabs. The control border is handled in Paint.
                        if (!isLast)
                        {
                            int right = bounds.Right - 1;
                            int top = bounds.Top + 4;
                            int bottom = bounds.Bottom - 5;
                            if (bottom <= top)
                            {
                                top = bounds.Top;
                                bottom = bounds.Bottom - 1;
                            }

                            e.Graphics.DrawLine(borderPen, right, top, right, bottom);
                        }
                    }

                    string label = tab.TabPages[e.Index].Text;
                    Rectangle textBounds = e.Bounds;
                    textBounds.Inflate(-6, 0);
                    TextRenderer.DrawText(
                        e.Graphics,
                        label,
                        tab.Font,
                        textBounds,
                        text,
                        TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter | TextFormatFlags.EndEllipsis);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to owner-draw TabControl tab.", ex);
                }
            };

            control.Paint += (sender, e) =>
            {
                try
                {
                    var tab = (TabControl)sender;
                    if (tab.TabPages.Count == 0)
                    {
                        return;
                    }

                    int stripHeight = tab.DisplayRectangle.Top;
                    if (stripHeight <= 0)
                    {
                        return;
                    }

                    Color selectedBack = palette.IsDark ? palette.ControlBackground : palette.WindowBackground;
                    Color unselectedBack = palette.IsDark ? palette.WindowBackground : palette.ControlBackground;

                    Rectangle stripRect = new Rectangle(0, 0, tab.ClientSize.Width, stripHeight);
                    using (var brush = new SolidBrush(unselectedBack))
                    using (var region = new Region(stripRect))
                    {
                        for (int i = 0; i < tab.TabPages.Count; i++)
                        {
                            Rectangle rect = tab.GetTabRect(i);
                            if (rect.IsEmpty)
                            {
                                continue;
                            }

                            region.Exclude(rect);
                        }

                        e.Graphics.FillRegion(brush, region);
                    }

                    // Extend the last tab into the remaining strip area so the strip always looks filled.
                    int lastIndex = tab.TabPages.Count - 1;
                    Rectangle lastRect = tab.GetTabRect(lastIndex);
                    if (!lastRect.IsEmpty && lastRect.Right < tab.ClientSize.Width)
                    {
                        bool lastSelected = lastIndex == tab.SelectedIndex;
                        Color lastBack = lastSelected ? selectedBack : unselectedBack;
                        Rectangle extendRect = new Rectangle(
                            lastRect.Right,
                            lastRect.Top,
                            tab.ClientSize.Width - lastRect.Right,
                            lastRect.Height);

                        using (var extendBrush = new SolidBrush(lastBack))
                        {
                            e.Graphics.FillRectangle(extendBrush, extendRect);
                        }
                    }

                    using (var borderPen = new Pen(palette.Border))
                    {
                        Rectangle border = tab.ClientRectangle;
                        border.Width = Math.Max(0, border.Width - 1);
                        border.Height = Math.Max(0, border.Height - 1);
                        e.Graphics.DrawRectangle(borderPen, border);

                        // Bottom separator under the tab strip.
                        int y = stripHeight - 1;
                        if (y > 0)
                        {
                            Rectangle selectedRect = Rectangle.Empty;
                            if (tab.SelectedIndex >= 0 && tab.SelectedIndex < tab.TabPages.Count)
                            {
                                selectedRect = tab.GetTabRect(tab.SelectedIndex);
                            }

                            int xRight = tab.ClientSize.Width - 1;
                            if (selectedRect.IsEmpty)
                            {
                                e.Graphics.DrawLine(borderPen, 0, y, xRight, y);
                            }
                            else
                            {
                                int breakLeft = Math.Max(0, selectedRect.Left - 1);
                                int breakRight = selectedRect.Right;
                                if (tab.SelectedIndex == tab.TabPages.Count - 1)
                                {
                                    breakRight = tab.ClientSize.Width;
                                }

                                if (breakLeft > 0)
                                {
                                    e.Graphics.DrawLine(borderPen, 0, y, breakLeft, y);
                                }

                                if (breakRight < xRight)
                                {
                                    e.Graphics.DrawLine(borderPen, breakRight, y, xRight, y);
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to paint TabControl background strip.", ex);
                }
            };
        }

        private static void DisableWindowsTheme(Control control)
        {
            if (control == null || control.Handle == IntPtr.Zero)
            {
                return;
            }

            try
            {
                // Disable themed painting for this control to prevent bright UI artifacts in dark mode and to
                // make custom BackColor/OwnerDraw behave consistently.
                SetWindowTheme(control.Handle, string.Empty, string.Empty);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to disable Windows theme for control.", ex);
            }
        }

        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        private static extern int SetWindowTheme(IntPtr hWnd, string pszSubAppName, string pszSubIdList);

        private static void ApplyTabStretchToControlTree(Control root)
        {
            if (root == null)
            {
                return;
            }

            if (root is TabControl)
            {
                ApplyTabStretch((TabControl)root);
            }

            foreach (Control child in root.Controls)
            {
                ApplyTabStretchToControlTree(child);
            }
        }

        private static void ApplyTabStretch(TabControl control)
        {
            if (control == null)
            {
                return;
            }

            control.SizeMode = TabSizeMode.Fixed;
            ApplyTabSizing(control);

            object existing;
            if (TabStretchApplied.TryGetValue(control, out existing))
            {
                return;
            }

            TabStretchApplied.Add(control, new object());

            control.FontChanged += (s, e) =>
            {
                try
                {
                    ApplyTabSizing(control);
                    control.Invalidate(true);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to re-apply tab sizing on FontChanged.", ex);
                }
            };

            control.Resize += (s, e) =>
            {
                try
                {
                    ApplyTabSizing(control);
                    control.Invalidate(true);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to re-apply tab sizing on Resize.", ex);
                }
            };

            control.ControlAdded += (s, e) =>
            {
                try
                {
                    ApplyTabSizing(control);
                    control.Invalidate(true);
                }
                catch (Exception ex)
                {
                    DiagnosticsLogger.LogException(LogCategories.Core, "Failed to re-apply tab sizing on ControlAdded.", ex);
                }
            };
        }

        private static void ApplyTabSizing(TabControl control)
        {
            if (control == null || control.TabPages == null || control.TabPages.Count == 0)
            {
                return;
            }

            int tabCount = control.TabPages.Count;
            int availableWidth = control.ClientSize.Width;
            if (availableWidth <= 0)
            {
                availableWidth = control.Width;
            }

            // Stretch tabs to fill the full control width so there is no visible strip background.
            int fillWidth = 0;
            if (tabCount > 0 && availableWidth > 0)
            {
                int dpi = control.DeviceDpi > 0 ? control.DeviceDpi : 96;
                int safetyMargin = (int)Math.Round(16f * (dpi / 96f));

                // Keep a safety margin to avoid triggering scroll arrows because of rounding/internal tab padding.
                fillWidth = Math.Max(0, (availableWidth - safetyMargin) / tabCount);
            }

            // Prefer no scroll arrows over perfect text fit; use equal-width tabs that always fit the strip.
            int desiredWidth = fillWidth;
            if (desiredWidth <= 0)
            {
                return;
            }
            int desiredHeight = Math.Max(20, control.ItemSize.Height);

            if (control.ItemSize.Width != desiredWidth || control.ItemSize.Height != desiredHeight)
            {
                control.ItemSize = new Size(desiredWidth, desiredHeight);
            }
        }

        private static void TrySetDarkTitleBar(IntPtr handle, bool enabled)
        {
            if (handle == IntPtr.Zero)
            {
                return;
            }

            try
            {
                int useDark = enabled ? 1 : 0;

                // Windows 10 1809+ uses 19, Windows 10 1903+ uses 20. Try both.
                if (DwmSetWindowAttribute(handle, 20, ref useDark, sizeof(int)) != 0)
                {
                    DwmSetWindowAttribute(handle, 19, ref useDark, sizeof(int));
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to set dark title bar attribute.", ex);
            }
        }

        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int dwAttribute, ref int pvAttribute, int cbAttribute);
    }

    internal static class OutlookThemeDetector
    {
        internal static OfficeUiTheme? TryGetOfficeUiTheme()
        {
            try
            {
                return TryReadOfficeUiTheme();
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read Office UI theme from registry.", ex);
                return null;
            }
        }

        internal static bool IsDarkThemePreferred()
        {
            try
            {
                // 1) Office UI theme (if exposed in the registry)
                OfficeUiTheme? officeTheme = TryReadOfficeUiTheme();
                if (officeTheme.HasValue)
                {
                    return officeTheme.Value == OfficeUiTheme.Black || officeTheme.Value == OfficeUiTheme.DarkGray;
                }

                // 2) Fallback to Windows app theme ("AppsUseLightTheme" == 0 -> dark)
                int? appsUseLight = ReadRegistryDword(Registry.CurrentUser, "Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize", "AppsUseLightTheme");
                if (appsUseLight.HasValue)
                {
                    return appsUseLight.Value == 0;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed while detecting dark theme preference.", ex);
            }

            return false;
        }

        private static OfficeUiTheme? TryReadOfficeUiTheme()
        {
            string[] versions = new[] { "16.0", "15.0", "14.0" };
            string[] keyPaths = new[]
            {
                "Software\\Microsoft\\Office\\{0}\\Common\\General",
                "Software\\Microsoft\\Office\\{0}\\Common",
                "Software\\Microsoft\\Office\\{0}\\Outlook\\Options\\General",
                "Software\\Microsoft\\Office\\{0}\\Outlook\\Options",
                "Software\\Microsoft\\Office\\{0}\\Outlook\\Preferences"
            };

            string[] valueNames = new[] { "UI Theme", "UITheme", "OfficeTheme", "Theme" };

            foreach (string version in versions)
            {
                foreach (string keyPattern in keyPaths)
                {
                    string key = string.Format(keyPattern, version);

                    // Outlook sometimes stores theme details inside a JSON payload (MonarchInboxPrimingData).
                    OfficeUiTheme? jsonTheme = TryReadOfficeThemeFromOutlookJson(key);
                    if (jsonTheme.HasValue)
                    {
                        return jsonTheme.Value;
                    }

                    foreach (string name in valueNames)
                    {
                        object value = ReadRegistryValue(Registry.CurrentUser, key, name);
                        if (value == null)
                        {
                            continue;
                        }

                        OfficeUiTheme? parsed = ParseOfficeThemeValue(value);
                        if (parsed.HasValue)
                        {
                            return parsed.Value;
                        }
                    }
                }
            }

            return null;
        }

        private static OfficeUiTheme? TryReadOfficeThemeFromOutlookJson(string outlookOptionsKey)
        {
            if (string.IsNullOrWhiteSpace(outlookOptionsKey))
            {
                return null;
            }

            try
            {
                object raw = ReadRegistryValue(Registry.CurrentUser, outlookOptionsKey, "MonarchInboxPrimingData");
                string json = raw as string;
                if (string.IsNullOrWhiteSpace(json))
                {
                    return null;
                }

                // Example snippet:
                // "OfficeThemeInfo":{"Theme":"Use system setting","IsReallyDarkTheme":true}
                Match themeMatch = Regex.Match(json, "\"OfficeThemeInfo\"\\s*:\\s*\\{[^}]*\"Theme\"\\s*:\\s*\"(?<theme>[^\"]+)\"", RegexOptions.IgnoreCase);
                Match darkMatch = Regex.Match(json, "\"OfficeThemeInfo\"\\s*:\\s*\\{[^}]*\"IsReallyDarkTheme\"\\s*:\\s*(?<dark>true|false)", RegexOptions.IgnoreCase);

                string themeText = themeMatch.Success ? themeMatch.Groups["theme"].Value : string.Empty;
                bool? isReallyDark = null;
                if (darkMatch.Success)
                {
                    bool parsed;
                    if (bool.TryParse(darkMatch.Groups["dark"].Value, out parsed))
                    {
                        isReallyDark = parsed;
                    }
                }

                if (!string.IsNullOrEmpty(themeText))
                {
                    string normalized = themeText.Trim().ToLowerInvariant();
                    if (normalized.Contains("black"))
                    {
                        return OfficeUiTheme.Black;
                    }
                    if (normalized.Contains("dark gray") || normalized.Contains("dark grey"))
                    {
                        return OfficeUiTheme.DarkGray;
                    }
                    if (normalized.Contains("white"))
                    {
                        return OfficeUiTheme.White;
                    }
                    if (normalized.Contains("colorful") || normalized.Contains("colourful"))
                    {
                        return OfficeUiTheme.Colorful;
                    }
                    if (normalized.Contains("use system"))
                    {
                        if (isReallyDark.HasValue)
                        {
                            return isReallyDark.Value ? OfficeUiTheme.Black : OfficeUiTheme.White;
                        }

                        // Fall back to Windows app theme to resolve "Use system setting".
                        int? appsUseLight = ReadRegistryDword(Registry.CurrentUser, "Software\\Microsoft\\Windows\\CurrentVersion\\Themes\\Personalize", "AppsUseLightTheme");
                        if (appsUseLight.HasValue)
                        {
                            return appsUseLight.Value == 0 ? OfficeUiTheme.Black : OfficeUiTheme.White;
                        }
                    }
                }

                if (isReallyDark.HasValue)
                {
                    return isReallyDark.Value ? OfficeUiTheme.Black : OfficeUiTheme.White;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to parse Outlook theme JSON payload.", ex);
            }

            return null;
        }

        private static OfficeUiTheme? ParseOfficeThemeValue(object rawValue)
        {
            if (rawValue == null)
            {
                return null;
            }

            try
            {
                // Common Office mapping (Office 2016/2019/365):
                // 0 = Colorful, 1 = Dark Gray, 2 = Black, 3 = White.
                if (rawValue is int)
                {
                    int numeric = (int)rawValue;
                    if (numeric == 0) return OfficeUiTheme.Colorful;
                    if (numeric == 1) return OfficeUiTheme.DarkGray;
                    if (numeric == 2) return OfficeUiTheme.Black;
                    if (numeric == 3) return OfficeUiTheme.White;
                }

                string text = rawValue as string;
                if (!string.IsNullOrEmpty(text))
                {
                    string normalized = text.Trim().ToLowerInvariant();
                    if (normalized.Contains("dark gray") || normalized.Contains("dark grey"))
                    {
                        return OfficeUiTheme.DarkGray;
                    }
                    if (normalized.Contains("black") || normalized.Contains("dark"))
                    {
                        return OfficeUiTheme.Black;
                    }
                    if (normalized.Contains("colorful") || normalized.Contains("colourful"))
                    {
                        return OfficeUiTheme.Colorful;
                    }
                    if (normalized.Contains("white"))
                    {
                        return OfficeUiTheme.White;
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to parse Office UI theme value.", ex);
            }

            return null;
        }

        private static int? ReadRegistryDword(RegistryKey root, string subKey, string valueName)
        {
            object value = ReadRegistryValue(root, subKey, valueName);
            if (value == null)
            {
                return null;
            }

            try
            {
                if (value is int)
                {
                    return (int)value;
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read registry DWORD value '" + valueName + "' from '" + subKey + "'.", ex);
            }

            return null;
        }

        private static object ReadRegistryValue(RegistryKey root, string subKey, string valueName)
        {
            if (root == null || string.IsNullOrWhiteSpace(subKey) || string.IsNullOrWhiteSpace(valueName))
            {
                return null;
            }

            try
            {
                using (var key = root.OpenSubKey(subKey))
                {
                    if (key == null)
                    {
                        return null;
                    }

                    return key.GetValue(valueName);
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "Failed to read registry value '" + valueName + "' from '" + subKey + "'.", ex);
                return null;
            }
        }
    }
}
