// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    // Selection stays local until confirmation; filtering never discards checked resources.
    internal sealed class CalendarResourceForm : ScaledForm
    {
        private readonly List<CalendarResource> _resources;
        private readonly HashSet<string> _selected;
        private readonly TextBox _search = new TextBox();
        private readonly ResourceListView _list = new ResourceListView();
        private readonly TableLayoutPanel _details = new TableLayoutPanel();
        private readonly Label _empty = new Label();
        private readonly Button _apply = new Button();
        private readonly ComboBox _grouping = new ComboBox();
        private readonly CheckBox _onlyAvailable = new CheckBox();
        private readonly ComboBox _buildingFilter = new ComboBox();
        private readonly ComboBox _floorFilter = new ComboBox();
        private readonly NumericUpDown _capacityFilter = new NumericUpDown();
        private readonly Button _featuresFilter = new Button();
        private readonly ContextMenuStrip _featuresMenu = new ContextMenuStrip();
        private readonly HashSet<string> _selectedFeatures = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly Button _expandAll = new Button();
        private readonly Button _clearFilters = new Button();
        private readonly Dictionary<string, bool> _collapsedFloors = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        private bool _groupsExpanded = true;
        private bool _updatingFilters;
        private Action<bool> _saveGroupsExpanded;
        private Font _floorFont;
        private readonly Label _interval = new Label();
        private readonly ContextMenuStrip _resourceMenu = new ContextMenuStrip();
        private readonly ToolTip _calendarTip = new ToolTip();
        private readonly System.Windows.Forms.Timer _intervalTimer = new System.Windows.Forms.Timer { Interval = 750 };
        private readonly Dictionary<string, CalendarResourceAvailabilityState> _availability = new Dictionary<string, CalendarResourceAvailabilityState>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, bool> _collapsed = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        private CalDavFreeBusyService _availabilityService;
        private CalendarResourceAvailabilityForm _dayView;
        private bool _openingDayView;
        private CalendarResourceTimeWindow _window;
        private Func<CalendarResourceTimeWindow> _readWindow;
        private Action<string> _saveGrouping;
        private CancellationTokenSource _availabilityCancellation;
        private int _availabilityGeneration;
        private bool _closing;
        private readonly Panel _policyWarningPanel = new Panel();
        private readonly Label _policyWarningTitleLabel = new Label();
        private readonly Label _policyWarningTextLabel = new Label();
        private readonly LinkLabel _policyWarningLinkLabel = new LinkLabel();
        private readonly List<Bitmap> _iconImages = new List<Bitmap>();
        private bool _updating;
        private bool _layoutPolicyWarning;
        private bool _initialPolicyWarningSizeApplied;
        private int[] _preferredColumnWidths, _minimumColumnWidths;
        private Font _columnFont;
        private int _columnDpi;
        private bool _sizingColumns;
        private bool _layoutDetails;

        private sealed class FloorGroup
        {
            internal string Key;
            internal bool Collapsed;
        }

        private sealed class ResourceListView : ListView
        {
            internal bool ChangingHandle;
            private bool _groupHeaderPressed;
            internal Action<ListViewItem, bool> ChangeFloor;
            internal Action<ListViewItem> OpenCalendar;
            internal Action HandleReady;

            internal bool IsCalendarHit(Point point)
            {
                ListViewHitTestInfo hit = HitTest(point);
                return hit.Item != null && hit.Item.Tag is CalendarResource && hit.SubItem != null && hit.Item.SubItems.IndexOf(hit.SubItem) == 5;
            }

            private bool HandleFloorKey(Keys key)
            {
                ListViewItem item = FocusedItem;
                var floor = item == null ? null : item.Tag as FloorGroup;
                if (floor == null || ChangeFloor == null || (key != Keys.Left && key != Keys.Right && key != Keys.Space && key != Keys.Enter)) { return false; }
                ChangeFloor(item, key == Keys.Left || (key != Keys.Right && !floor.Collapsed));
                return true;
            }

            // Native group headings ignore ListView.ForeColor under visual styles.
            // Repaint only headings; native items retain keyboard and accessibility support.
            [StructLayout(LayoutKind.Sequential)]
            private struct NotificationHeader
            {
                public IntPtr Window;
                public UIntPtr Id;
                public int Code;
            }

            [StructLayout(LayoutKind.Sequential)]
            private struct NativeRectangle
            {
                public int Left, Top, Right, Bottom;
            }

            [StructLayout(LayoutKind.Sequential)]
            private struct ItemChange
            {
                public NotificationHeader Header;
                public int Item, SubItem;
                public uint NewState, OldState, Changed;
                public Point Action;
                public IntPtr Data;
            }

            [StructLayout(LayoutKind.Sequential)]
            private struct NativeItem
            {
                public uint Mask;
                public int Item, SubItem;
                public uint State, StateMask;
                public IntPtr Text;
                public int TextLength, Image;
                public IntPtr Data;
                public int Indent, GroupId;
                public uint ColumnCount;
                public IntPtr Columns, ColumnFormats;
                public int Group;
            }

            [StructLayout(LayoutKind.Sequential)]
            private struct CustomDraw
            {
                public NotificationHeader Header;
                public uint Stage;
                public IntPtr DeviceContext;
                public NativeRectangle Bounds;
                public UIntPtr Item;
                public uint State;
                public IntPtr ItemData;
            }

            // Layout follows commctrl.h NMLVCUSTOMDRAW on both x86 and x64.
            [StructLayout(LayoutKind.Sequential)]
            private struct ListCustomDraw
            {
                public CustomDraw Draw;
                public int TextColor, TextBackground;
                public int SubItem;
                public uint ItemType;
                public int FaceColor, IconEffect, IconPhase, Part, PartState;
                public NativeRectangle TextBounds;
                public uint Align;
            }

            [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
            private struct NativeGroup
            {
                public uint Size, Mask;
                public IntPtr Header;
                public int HeaderLength;
                public IntPtr Footer;
                public int FooterLength, GroupId;
                public uint StateMask, State, Align;
                public IntPtr Subtitle;
                public uint SubtitleLength;
                public IntPtr Task;
                public uint TaskLength;
                public IntPtr DescriptionTop;
                public uint DescriptionTopLength;
                public IntPtr DescriptionBottom;
                public uint DescriptionBottomLength;
                public int TitleImage, ExtendedImage, FirstItem;
                public uint ItemCount;
                public IntPtr SubsetTitle;
                public uint SubsetTitleLength;
            }

            [DllImport("user32.dll", CharSet = CharSet.Unicode)]
            private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr parameter, ref NativeGroup group);

            [DllImport("user32.dll", CharSet = CharSet.Unicode)]
            private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr parameter, ref NativeRectangle rectangle);

            [DllImport("user32.dll", CharSet = CharSet.Unicode)]
            private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr parameter, IntPtr value);

            [DllImport("user32.dll", CharSet = CharSet.Unicode)]
            private static extern IntPtr SendMessage(IntPtr window, int message, IntPtr parameter, ref NativeItem item);

            internal void HideFloorCheckboxes()
            {
                if (!IsHandleCreated || ChangingHandle) { return; }
                foreach (ListViewItem item in Items)
                {
                    if (!(item.Tag is FloorGroup)) { continue; }
                    // StateImageIndex caches its value; native recreation can reintroduce the checkbox independently.
                    var state = new NativeItem { StateMask = 0xF000 };
                    SendMessage(Handle, 0x102B, new IntPtr(item.Index), ref state);
                }
            }

            private bool TryGetGroup(int index, out NativeGroup group)
            {
                group = new NativeGroup { Size = (uint)Marshal.SizeOf(typeof(NativeGroup)), Mask = 0x14, StateMask = 0x11 };
                return IsHandleCreated && index >= 0 && index < Groups.Count
                    && SendMessage(Handle, 0x1099, new IntPtr(index), ref group) != IntPtr.Zero;
            }

            private Rectangle GroupHeaderBounds(int groupId)
            {
                var rectangle = new NativeRectangle { Top = 1 }; // LVGGR_HEADER
                return SendMessage(Handle, 0x1062, new IntPtr(groupId), ref rectangle) != IntPtr.Zero
                    ? Rectangle.FromLTRB(rectangle.Left, rectangle.Top, rectangle.Right, rectangle.Bottom) : Rectangle.Empty;
            }

            private bool TryGetHeader(Point point, out NativeGroup group)
            {
                for (int index = 0; index < Groups.Count; index++)
                { if (TryGetGroup(index, out group) && GroupHeaderBounds(group.GroupId).Contains(point)) { return true; } }
                group = new NativeGroup();
                return false;
            }

            private void SetGroupCollapsed(NativeGroup group, bool collapsed)
            {
                Focus();
                if (FocusedItem != null) { FocusedItem.Focused = false; }
                NativeGroup previous;
                int focused = SendMessage(Handle, 0x105D, IntPtr.Zero, IntPtr.Zero).ToInt32();
                if (TryGetGroup(focused, out previous) && previous.GroupId != group.GroupId)
                {
                    previous.Mask = 4; previous.StateMask = 0x10; previous.State = 0;
                    SendMessage(Handle, 0x1093, new IntPtr(previous.GroupId), ref previous);
                }
                group.Mask = 4; group.StateMask = 0x11; group.State = 0x10U | (collapsed ? 1U : 0U);
                SendMessage(Handle, 0x1093, new IntPtr(group.GroupId), ref group);
                Invalidate();
            }

            private bool HandleGroupKey(Keys key)
            {
                if (!IsHandleCreated || FocusedItem != null || (key != Keys.Left && key != Keys.Right && key != Keys.Space && key != Keys.Enter)) { return false; }
                NativeGroup group;
                if (!TryGetGroup(SendMessage(Handle, 0x105D, IntPtr.Zero, IntPtr.Zero).ToInt32(), out group)) { return false; }
                SetGroupCollapsed(group, key == Keys.Left || (key != Keys.Right && (group.State & 1) == 0));
                return true;
            }

            protected override bool ProcessCmdKey(ref Message message, Keys keyData)
            { return HandleFloorKey(keyData) || HandleGroupKey(keyData) || base.ProcessCmdKey(ref message, keyData); }

            protected override bool ProcessDialogKey(Keys keyData)
            { return HandleFloorKey(keyData) || HandleGroupKey(keyData) || base.ProcessDialogKey(keyData); }

            protected override void OnMouseMove(MouseEventArgs e)
            {
                base.OnMouseMove(e);
                NativeGroup group;
                ListViewItem item = GetItemAt(e.X, e.Y);
                Cursor = TryGetHeader(e.Location, out group) || (item != null && item.Tag is FloorGroup) || IsCalendarHit(e.Location) ? Cursors.Hand : Cursors.Default;
            }

            protected override void OnMouseLeave(EventArgs e)
            { Cursor = Cursors.Default; base.OnMouseLeave(e); }

            internal void CaptureGroups(IDictionary<string, bool> collapsed)
            {
                if (!IsHandleCreated) { return; }
                for (int index = 0; index < Groups.Count; index++)
                {
                    NativeGroup group;
                    if (TryGetGroup(index, out group))
                    { collapsed[Groups[index].Name] = (group.State & 1) != 0; }
                }
            }

            internal void RestoreGroups(IDictionary<string, bool> collapsed)
            {
                if (!IsHandleCreated) { return; }
                for (int index = 0; index < Groups.Count; index++)
                {
                    NativeGroup group;
                    if (!TryGetGroup(index, out group)) { continue; }
                    bool isCollapsed;
                    group.Mask = 4;
                    group.StateMask = 9;
                    group.State = 8U | (collapsed.TryGetValue(Groups[index].Name, out isCollapsed) && isCollapsed ? 1U : 0U);
                    SendMessage(Handle, 0x1093, new IntPtr(group.GroupId), ref group);
                }
            }

            protected override void WndProc(ref Message message)
            {
                // Headers have no checkbox state, including changes made by native accessibility actions.
                if (!ChangingHandle && message.Msg == 0x204E && message.LParam != IntPtr.Zero
                    && ((NotificationHeader)Marshal.PtrToStructure(message.LParam, typeof(NotificationHeader))).Code == -100)
                {
                    var change = (ItemChange)Marshal.PtrToStructure(message.LParam, typeof(ItemChange));
                    if (change.Item >= 0 && change.Item < Items.Count && Items[change.Item].Tag is FloorGroup
                        && (change.Changed & 8) != 0)
                    {
                        // WinForms interprets a missing state image as checked; headers must bypass that mapping.
                        message.Result = (change.NewState & 0xF000) == 0 ? IntPtr.Zero : new IntPtr(1);
                        return;
                    }
                }
                if (message.Msg == 0x201 || message.Msg == 0x203) // WM_LBUTTONDOWN / WM_LBUTTONDBLCLK
                {
                    _groupHeaderPressed = false;
                    long coordinates = message.LParam.ToInt64();
                    Point point = new Point((short)coordinates, (short)(coordinates >> 16));
                    NativeGroup hitGroup;
                    if (TryGetHeader(point, out hitGroup))
                    {
                        _groupHeaderPressed = true;
                        if (message.Msg == 0x201) { SetGroupCollapsed(hitGroup, (hitGroup.State & 1) == 0); }
                        return;
                    }
                    ListViewItem item = HitTest(point).Item;
                    if (IsCalendarHit(point) && OpenCalendar != null)
                    {
                        _groupHeaderPressed = true;
                        if (message.Msg == 0x201) { OpenCalendar(item); }
                        return;
                    }
                    var floor = item == null ? null : item.Tag as FloorGroup;
                    if (floor != null && ChangeFloor != null)
                    {
                        _groupHeaderPressed = true;
                        Focus();
                        if (message.Msg == 0x201) { ChangeFloor(item, !floor.Collapsed); }
                        return;
                    }
                }
                if (message.Msg == 0x202 && _groupHeaderPressed) { _groupHeaderPressed = false; return; }
                if (message.Msg == 0x100 && (HandleFloorKey((Keys)message.WParam.ToInt32() | ModifierKeys) || HandleGroupKey((Keys)message.WParam.ToInt32() | ModifierKeys))) { return; }
                base.WndProc(ref message);
                const int reflectedNotify = 0x204E;
                const int customDraw = -12;
                const uint prepaint = 1, postpaint = 2, group = 1;
                if (message.Msg != reflectedNotify || message.LParam == IntPtr.Zero || SystemInformation.HighContrast) { return; }
                var header = (NotificationHeader)Marshal.PtrToStructure(message.LParam, typeof(NotificationHeader));
                if (header.Code != customDraw) { return; }
                var draw = (ListCustomDraw)Marshal.PtrToStructure(message.LParam, typeof(ListCustomDraw));
                if (draw.Draw.Stage == prepaint)
                {
                    message.Result = new IntPtr(message.Result.ToInt64() | 0x20); // CDRF_NOTIFYITEMDRAW
                    if (draw.ItemType == group) { message.Result = new IntPtr(message.Result.ToInt64() | 0x10); } // CDRF_NOTIFYPOSTPAINT
                }
                else if (draw.Draw.Stage == postpaint && draw.ItemType == group)
                {
                    PaintGroupHeader(draw);
                }
            }

            private void PaintGroupHeader(ListCustomDraw draw)
            {
                Rectangle bounds = Rectangle.FromLTRB(draw.TextBounds.Left, draw.TextBounds.Top, draw.TextBounds.Right, draw.TextBounds.Bottom);
                if (bounds.Width <= 0 || bounds.Height <= 0 || draw.Draw.DeviceContext == IntPtr.Zero) { return; }
                IntPtr buffer = Marshal.AllocHGlobal(1024);
                try
                {
                    var group = new NativeGroup { Size = (uint)Marshal.SizeOf(typeof(NativeGroup)), Mask = 5, StateMask = 0x11, Header = buffer, HeaderLength = 512 };
                    if (SendMessage(Handle, 0x1095, new IntPtr(unchecked((long)draw.Draw.Item.ToUInt64())), ref group).ToInt64() == -1) { return; }
                    string title = Marshal.PtrToStringUni(group.Header) ?? string.Empty;
                    Rectangle headerBounds = GroupHeaderBounds(unchecked((int)draw.Draw.Item.ToUInt64()));
                    if (!headerBounds.IsEmpty) { bounds = headerBounds; }
                    using (Graphics graphics = Graphics.FromHdc(draw.Draw.DeviceContext))
                    using (var background = new SolidBrush(BackColor.GetBrightness() < 0.5f ? ControlPaint.Light(BackColor, 0.12f) : ControlPaint.Dark(BackColor, 0.05f)))
                    using (var foreground = new SolidBrush(ForeColor))
                    using (var font = new Font(Font, FontStyle.Bold))
                    using (var line = new Pen(ForeColor, 1))
                    {
                        // The Windows theme ignores the group text colour supplied in NMLVCUSTOMDRAW.
                        graphics.FillRectangle(background, bounds);
                        int padding = Math.Max(6, DeviceDpi * 10 / 96);
                        bounds.Inflate(-padding, 0);
                        int arrowSize = Math.Max(5, DeviceDpi * 6 / 96);
                        int arrowX = bounds.Left + arrowSize;
                        int arrowY = bounds.Top + bounds.Height / 2;
                        Point[] arrow = (group.State & 1) != 0
                            ? new[] { new Point(arrowX - arrowSize, arrowY - arrowSize), new Point(arrowX - arrowSize, arrowY + arrowSize), new Point(arrowX + arrowSize, arrowY) }
                            : new[] { new Point(arrowX - arrowSize, arrowY - arrowSize / 2), new Point(arrowX + arrowSize, arrowY - arrowSize / 2), new Point(arrowX, arrowY + arrowSize) };
                        graphics.FillPolygon(foreground, arrow);
                        int textOffset = arrowSize * 2 + padding;
                        bounds.X += textOffset;
                        bounds.Width = Math.Max(0, bounds.Width - textOffset);
                        TextRenderer.DrawText(graphics, title, font, bounds, ForeColor,
                            TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding | TextFormatFlags.EndEllipsis);
                        int lineStart = bounds.Left + TextRenderer.MeasureText(graphics, title, font, Size.Empty,
                            TextFormatFlags.NoPadding | TextFormatFlags.NoPrefix).Width + padding;
                        if (lineStart < bounds.Right) { graphics.DrawLine(line, lineStart, bounds.Top + bounds.Height / 2, bounds.Right, bounds.Top + bounds.Height / 2); }
                        if (Focused && (group.State & 0x10) != 0 && !headerBounds.IsEmpty)
                        { ControlPaint.DrawFocusRectangle(graphics, Rectangle.Inflate(headerBounds, -2, -2), ForeColor, BackColor); }
                    }
                }
                finally { Marshal.FreeHGlobal(buffer); }
            }

            protected override void CreateHandle()
            {
                ChangingHandle = true;
                try { base.CreateHandle(); }
                finally { ChangingHandle = false; }
                // Native checkbox styles restore default state images after handle recreation returns.
                BeginInvoke((MethodInvoker)delegate { if (!IsDisposed && IsHandleCreated && HandleReady != null) { HandleReady(); } });
            }
            protected override void DestroyHandle()
            {
                ChangingHandle = true;
                try { base.DestroyHandle(); }
                finally { ChangingHandle = false; }
            }
        }

        internal CalendarResourceForm(IEnumerable<CalendarResource> resources, IEnumerable<string> selected, bool ifbEnabled, BackendPolicyStatus policyStatus, string baseUrl)
        {
            _resources = resources.OrderBy(r => !r.IsRoom).ThenBy(r => r.DisplayName, StringComparer.CurrentCultureIgnoreCase).ToList();
            _selected = new HashSet<string>(selected, StringComparer.OrdinalIgnoreCase);
            Text = Strings.ResourceDialogTitle;
            Icon = BrandingAssets.GetAppIcon(32);
            Font = new Font("Segoe UI", 9F);
            AutoScaleMode = AutoScaleMode.Dpi;
            MinimumSize = new Size(1000, 760);
            ClientSize = new Size(1240, 820);
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            MinimizeBox = false;
            MaximizeBox = true;

            var layout = new TableLayoutPanel { Dock = DockStyle.Fill, Padding = new Padding(16), ColumnCount = 2, RowCount = 7 };
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 69));
            layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 31));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            WarningPanelUiHelper.Initialize(_policyWarningPanel, _policyWarningTitleLabel,
                _policyWarningTextLabel, _policyWarningLinkLabel, "\u26a0 " + Strings.PolicyWarningTitle,
                string.Empty, Strings.PolicyWarningAdminLinkLabel);
            _policyWarningPanel.Dock = DockStyle.Top;
            _policyWarningPanel.Margin = new Padding(0, 0, 0, 10);
            _policyWarningPanel.Layout += delegate { LayoutPolicyWarningPanel(); };
            _policyWarningPanel.VisibleChanged += delegate { LayoutPolicyWarningPanel(); };
            _policyWarningLinkLabel.LinkClicked += (s, e) =>
                PolicyUiHelper.OpenLicenseAdministration(_policyWarningLinkLabel, LogCategories.Ifb);
            layout.Controls.Add(_policyWarningPanel, 0, 0);
            layout.SetColumnSpan(_policyWarningPanel, 2);
            var selectionHint = new Label { Text = Strings.ResourceSelectionHint, AutoSize = true, Margin = new Padding(0, 0, 0, 10) };
            layout.Controls.Add(selectionHint, 0, 1);
            layout.SetColumnSpan(selectionHint, 2);
            _search.MinimumSize = new Size(180, 0);
            _search.AccessibleName = Strings.ResourceSearchLabel;
            _search.TextChanged += delegate { RebuildList(); };

            _interval.AutoSize = true;
            _interval.Margin = new Padding(0, 6, 0, 0);
            _interval.UseMnemonic = false;
            layout.Controls.Add(_interval, 0, 2);
            layout.SetColumnSpan(_interval, 2);
            ConfigureMetadataFilter(_buildingFilter, Strings.ResourceBuildingLabel, _resources.Select(r => r.BuildingName));
            ConfigureMetadataFilter(_floorFilter, Strings.ResourceFloorLabel, _resources.Select(r => r.Floor));
            _buildingFilter.SelectedIndexChanged += delegate
            {
                if (_updatingFilters) { return; }
                UpdateFloorFilter();
                RebuildList();
            };
            _floorFilter.SelectedIndexChanged += delegate { if (!_updatingFilters) { RebuildList(); } };
            _capacityFilter.Minimum = 0;
            _capacityFilter.Maximum = int.MaxValue;
            _capacityFilter.Width = 95;
            _capacityFilter.AccessibleName = Strings.ResourceMinimumCapacity;
            _capacityFilter.ValueChanged += delegate { if (!_updatingFilters) { RebuildList(); } };
            _featuresFilter.AutoSize = true;
            _featuresFilter.MinimumSize = new Size(145, 26);
            _featuresFilter.AccessibleName = Strings.ResourceFeaturesFilter;
            _featuresFilter.Click += delegate { _featuresMenu.Show(_featuresFilter, new Point(0, _featuresFilter.Height)); };
            _featuresMenu.ShowCheckMargin = true;
            _featuresMenu.ShowImageMargin = false;
            _featuresMenu.Closing += (s, e) => { if (e.CloseReason == ToolStripDropDownCloseReason.ItemClicked) { e.Cancel = true; } };
            foreach (string feature in _resources.SelectMany(FeatureValues).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(v => v, StringComparer.CurrentCultureIgnoreCase))
            {
                var option = new ToolStripMenuItem(feature) { CheckOnClick = true };
                option.CheckedChanged += delegate
                {
                    if (option.Checked) { _selectedFeatures.Add(feature); } else { _selectedFeatures.Remove(feature); }
                    UpdateFeaturesLabel();
                    if (!_updatingFilters) { RebuildList(); }
                };
                _featuresMenu.Items.Add(option);
            }
            _featuresFilter.Enabled = _featuresMenu.Items.Count > 0;
            UpdateFeaturesLabel();
            _onlyAvailable.Text = Strings.ResourceOnlyAvailable;
            _onlyAvailable.AutoSize = true;
            _onlyAvailable.CheckedChanged += delegate { RebuildList(); };
            _clearFilters.Text = Strings.ResourceClearFilters;
            _clearFilters.AutoSize = true;
            _clearFilters.Click += delegate { ClearFilters(); };
            var metadataFilters = FilterRow(_onlyAvailable, _clearFilters,
                FilterField(Strings.ResourceBuildingLabel, _buildingFilter),
                FilterField(Strings.ResourceFloorLabel, _floorFilter),
                FilterField(Strings.ResourceMinimumCapacity, _capacityFilter),
                FilterField(Strings.ResourceFeaturesFilter, _featuresFilter),
                FilterField(null, _onlyAvailable), FilterField(null, _clearFilters));
            metadataFilters.Margin = new Padding(0, 8, 0, 8);
            layout.Controls.Add(metadataFilters, 0, 3);

            var content = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, Margin = new Padding(0, 8, 0, 8) };
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 69));
            content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 31));
            var listPanel = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 1, Margin = Padding.Empty };
            listPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            listPanel.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            _grouping.DropDownStyle = ComboBoxStyle.DropDownList;
            _grouping.Width = _buildingFilter.Width;
            _grouping.AccessibleName = Strings.ResourceGroupByLabel;
            _grouping.Items.Add(new GroupingOption("none", Strings.ResourceGroupingNone));
            if (_resources.Any(r => !string.IsNullOrWhiteSpace(r.BuildingName))) { _grouping.Items.Add(new GroupingOption("building", Strings.ResourceBuildingLabel)); }
            if (_resources.Any(r => !string.IsNullOrWhiteSpace(r.Floor))) { _grouping.Items.Add(new GroupingOption("floor", Strings.ResourceFloorLabel)); }
            if (_resources.Any(r => r.IsRoom && !string.IsNullOrWhiteSpace(r.RoomType))) { _grouping.Items.Add(new GroupingOption("room-type", Strings.ResourceRoomTypeLabel)); }
            _grouping.SelectedIndex = 0;
            _grouping.SelectedIndexChanged += delegate
            {
                RebuildList();
                if (_saveGrouping != null) { _saveGrouping(((GroupingOption)_grouping.SelectedItem).Key); }
            };
            _expandAll.AutoSize = true;
            _expandAll.Text = Strings.ResourceCollapseAll;
            _expandAll.Click += delegate
            {
                SetGroupsExpanded(!_groupsExpanded);
                if (_saveGroupsExpanded != null) { _saveGroupsExpanded(_groupsExpanded); }
            };
            var filters = FilterRow(_search, _expandAll, FilterField(Strings.ResourceGroupByLabel, _grouping),
                FilterField(Strings.ResourceSearchLabel, _search), FilterField(null, _expandAll));
            filters.Margin = Padding.Empty;
            layout.Controls.Add(filters, 0, 4);
            _list.Dock = DockStyle.Fill;
            _list.View = View.Details;
            _list.CheckBoxes = true;
            _list.FullRowSelect = true;
            _list.HideSelection = false;
            _list.MultiSelect = false;
            _list.ShowItemToolTips = true;
            _list.SmallImageList = CreateResourceIcons(ScaleLogical(20));
            _list.Columns.Add(Strings.ResourceNameLabel);
            _list.Columns.Add(Strings.ResourceRoomTypeLabel);
            _list.Columns.Add(Strings.ResourceBuildingLabel);
            _list.Columns.Add(new ColumnHeader { Text = Strings.ResourceCapacityLabel, TextAlign = HorizontalAlignment.Left });
            _list.Columns.Add(Strings.ResourceAvailabilityLabel);
            _list.Columns.Add(new ColumnHeader { Text = "📅", TextAlign = HorizontalAlignment.Center });
            _list.OpenCalendar = async item =>
            {
                _list.SelectedItems.Clear(); item.Selected = true; item.Focused = true;
                await ShowAvailabilityAsync();
            };
            _list.MouseMove += (s, e) =>
            {
                string tip = _list.IsCalendarHit(e.Location) ? Strings.ResourceShowAvailability : string.Empty;
                if (_calendarTip.GetToolTip(_list) != tip) { _calendarTip.SetToolTip(_list, tip); }
            };
            _list.HandleReady = delegate { _list.RestoreGroups(_collapsed); _list.HideFloorCheckboxes(); UpdateColumnWidths(); };
            _list.Resize += delegate { UpdateColumnWidths(); };
            _list.FontChanged += delegate { UpdateFloorFont(); UpdateColumnWidths(); };
            _list.DpiChangedAfterParent += delegate { UpdateColumnWidths(); };
            _list.ItemCheck += OnItemCheck;
            _list.SelectedIndexChanged += delegate { ShowDetails(SelectedResource); };
            _list.ChangeFloor = (item, collapsed) =>
            {
                string key = ((FloorGroup)item.Tag).Key;
                _collapsedFloors[key] = collapsed;
                RebuildList();
                ListViewItem header = _list.Items.Cast<ListViewItem>().FirstOrDefault(i => i.Tag is FloorGroup && ((FloorGroup)i.Tag).Key == key);
                if (header != null)
                {
                    _list.SelectedItems.Clear(); header.Selected = true; header.Focused = true; header.EnsureVisible();
                }
            };
            var showAvailability = new ToolStripMenuItem(Strings.ResourceShowAvailability);
            showAvailability.Click += async delegate { await ShowAvailabilityAsync(); };
            _resourceMenu.Items.Add(showAvailability);
            _resourceMenu.Opening += (s, e) => { e.Cancel = _availabilityService == null || _window == null || SelectedResource == null; };
            _list.ContextMenuStrip = _resourceMenu;
            _list.MouseDown += (s, e) =>
            {
                if (e.Button != MouseButtons.Right) { return; }
                ListViewItem item = _list.GetItemAt(e.X, e.Y);
                if (item == null) { _list.SelectedItems.Clear(); return; }
                item.Selected = true;
                item.Focused = true;
            };
            listPanel.Controls.Add(_list, 0, 0);
            content.Controls.Add(listPanel, 0, 0);
            _details.Dock = DockStyle.Fill;
            _details.AutoScroll = true;
            _details.ColumnCount = 2;
            _details.Padding = new Padding(12, 0, 0, 0);
            _details.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute));
            _details.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            _details.Resize += delegate { LayoutDetails(); };
            _details.FontChanged += delegate { LayoutDetails(); };
            _details.DpiChangedAfterParent += delegate { LayoutDetails(); };
            content.Controls.Add(_details, 1, 0);
            layout.Controls.Add(content, 0, 5);
            layout.SetColumnSpan(content, 2);
            var footer = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Top, ColumnCount = 2, RowCount = 1, Margin = Padding.Empty };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            footer.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            var notices = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Top, ColumnCount = 1, Margin = new Padding(0, 0, 12, 0) };
            notices.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            _empty.AutoSize = true;
            _empty.Dock = DockStyle.Top;
            _empty.Margin = Padding.Empty;
            _empty.Text = Strings.ResourceNoResults;
            notices.Controls.Add(_empty);
            notices.Controls.Add(new Label { Text = Strings.ResourceBookingHint, AutoSize = true, Dock = DockStyle.Top, Margin = new Padding(0, 3, 0, 3) });
            if (!ifbEnabled)
            {
                notices.Controls.Add(new Label { Text = Strings.ResourceIfbDisabledHint, AutoSize = true, Dock = DockStyle.Top, Margin = new Padding(0, 0, 0, 3) });
            }
            footer.Controls.Add(notices, 0, 0);
            var buttons = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Anchor = AnchorStyles.Bottom | AnchorStyles.Right, FlowDirection = FlowDirection.RightToLeft, WrapContents = false, Margin = Padding.Empty };
            var cancel = new Button { Text = Strings.ButtonCancel, DialogResult = DialogResult.Cancel, AutoSize = true, MinimumSize = new Size(100, 32) };
            _apply.Text = Strings.ResourceApplyButton;
            _apply.AutoSize = true;
            _apply.MinimumSize = new Size(100, 32);
            _apply.DialogResult = DialogResult.OK;
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(_apply);
            footer.Controls.Add(buttons, 1, 0);
            layout.Controls.Add(footer, 0, 6);
            layout.SetColumnSpan(footer, 2);
            Controls.Add(layout);
            BrandedHeader.AttachToParent(new BrandedHeader(), Controls, 48);
            AcceptButton = _apply;
            CancelButton = cancel;
            UiThemeManager.ApplyToForm(this);
            RebuildList();
            PolicyUiHelper.ApplyPolicyWarningState(policyStatus, _policyWarningPanel,
                _policyWarningTextLabel, _policyWarningTitleLabel, _policyWarningLinkLabel, baseUrl);
        }

        protected override async void OnShown(EventArgs e)
        {
            base.OnShown(e);
            if (_closing || IsDisposed || Disposing) { return; }
            LayoutPolicyWarningPanel();
            AdjustInitialSizeForPolicyWarning(Screen.FromControl(this).WorkingArea);
            if (_availabilityService != null)
            {
                _intervalTimer.Start();
                await RefreshTimeWindowAsync();
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            if (!e.Cancel) { StopAvailability(); }
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            StopAvailability();
            base.OnFormClosed(e);
        }

        internal void ConfigureAvailability(CalDavFreeBusyService service, Func<CalendarResourceTimeWindow> readWindow,
            string grouping, Action<string> saveGrouping, bool groupsExpanded, Action<bool> saveGroupsExpanded)
        {
            _availabilityService = service;
            _readWindow = readWindow;
            foreach (GroupingOption option in _grouping.Items)
            {
                if (option.Key == grouping) { _grouping.SelectedItem = option; break; }
            }
            _saveGrouping = saveGrouping;
            SetGroupsExpanded(groupsExpanded);
            _saveGroupsExpanded = saveGroupsExpanded;
            _intervalTimer.Tick += async delegate { await RefreshTimeWindowAsync(); };
        }

        private sealed class FilterRowPanel : FlowLayoutPanel
        {
            private bool _sizingHeight;
            private bool _parentLayoutPending;

            protected override void OnLayout(LayoutEventArgs e)
            {
                base.OnLayout(e);
                if (_sizingHeight || Controls.Count == 0) { return; }
                _sizingHeight = true;
                try
                {
                    // Measure after native wrapping, not its cached pre-resize preferred height.
                    int height = Controls.Cast<Control>().Max(control => control.Bottom + control.Margin.Bottom) + Padding.Bottom;
                    var parent = Parent as TableLayoutPanel;
                    int rowIndex = parent == null ? -1 : parent.GetRow(this);
                    if (rowIndex >= 0 && rowIndex < parent.RowStyles.Count)
                    {
                        RowStyle style = parent.RowStyles[rowIndex];
                        style.SizeType = SizeType.Absolute;
                        int rowHeight = height + Margin.Vertical;
                        if (style.Height != rowHeight)
                        {
                            style.Height = rowHeight;
                            if (IsHandleCreated && !_parentLayoutPending)
                            {
                                _parentLayoutPending = true;
                                BeginInvoke(new Action(delegate
                                {
                                    _parentLayoutPending = false;
                                    if (!IsDisposed && !Disposing && !parent.IsDisposed && !parent.Disposing) { parent.PerformLayout(); }
                                }));
                            }
                        }
                    }
                    Height = height;
                }
                finally { _sizingHeight = false; }
            }
        }

        private FlowLayoutPanel FilterRow(Control flexibleInput, Button action, params TableLayoutPanel[] fields)
        {
            var row = new FilterRowPanel { Dock = DockStyle.Top, WrapContents = true };
            var flexibleField = (TableLayoutPanel)flexibleInput.Parent;
            flexibleField.AutoSize = false;
            flexibleField.ColumnStyles[0] = new ColumnStyle(SizeType.Percent, 100);
            flexibleInput.Anchor = flexibleInput is TextBox ? AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right : AnchorStyles.Right;
            action.AutoSize = false;
            var actionField = (TableLayoutPanel)action.Parent;
            actionField.AutoSize = false;
            actionField.ColumnStyles[0] = new ColumnStyle(SizeType.Percent, 100);
            actionField.Margin = Padding.Empty;
            action.Anchor = AnchorStyles.Right;
            bool sizing = false;
            row.Layout += delegate
            {
                if (sizing || row.IsDisposed || row.Disposing) { return; }
                sizing = true;
                try
                {
                    var buildingField = _buildingFilter.Parent as TableLayoutPanel;
                    var groupingField = _grouping.Parent as TableLayoutPanel;
                    if (buildingField != null && groupingField != null)
                    {
                        int firstColumnWidth = Math.Max(Math.Max(_buildingFilter.Width, _grouping.Width),
                            Math.Max(buildingField.GetControlFromPosition(0, 0).PreferredSize.Width,
                                groupingField.GetControlFromPosition(0, 0).PreferredSize.Width));
                        buildingField.MinimumSize = new Size(firstColumnWidth, 0);
                        groupingField.MinimumSize = new Size(firstColumnWidth, 0);
                    }
                    int expandTextWidth = Math.Max(TextRenderer.MeasureText(Strings.ResourceCollapseAll, _expandAll.Font).Width,
                        TextRenderer.MeasureText(Strings.ResourceExpandAll, _expandAll.Font).Width);
                    int actionWidth = Math.Max(TextRenderer.MeasureText(_clearFilters.Text, _clearFilters.Font).Width,
                        expandTextWidth) + Math.Max(_clearFilters.Padding.Horizontal, _expandAll.Padding.Horizontal)
                        + DpiScaling.ScaleLogical(this, 20);
                    _clearFilters.Size = new Size(actionWidth, _clearFilters.PreferredSize.Height);
                    _expandAll.Size = new Size(actionWidth, _expandAll.PreferredSize.Height);
                    int captionHeight = 0, inputHeight = 0;
                    foreach (TableLayoutPanel field in fields)
                    {
                        Control caption = field.GetControlFromPosition(0, 0);
                        if (caption != null) { captionHeight = Math.Max(captionHeight, caption.PreferredSize.Height + caption.Margin.Vertical); }
                        Control input = field.GetControlFromPosition(0, 1);
                        // Native input heights can differ from their reported preferred height.
                        inputHeight = Math.Max(inputHeight, input is ComboBox || input is TextBox ? input.Height : input.PreferredSize.Height);
                    }
                    foreach (TableLayoutPanel field in fields)
                    {
                        field.RowStyles[0].Height = captionHeight;
                        field.RowStyles[1].Height = inputHeight;
                        Control input = field.GetControlFromPosition(0, 1);
                        if (input is ComboBox || input is TextBox) { input.Margin = new Padding(0, Math.Max(0, (inputHeight - input.Height) / 2), 0, 0); }
                    }
                    int availableWidth = row.ClientSize.Width - row.Padding.Horizontal;
                    int lineWidth = 0;
                    TableLayoutPanel precedingField = null;
                    foreach (TableLayoutPanel field in fields.TakeWhile(field => field != flexibleField))
                    {
                        row.SetFlowBreak(field, false);
                        int width = field.GetPreferredSize(Size.Empty).Width + field.Margin.Horizontal;
                        if (lineWidth + width > availableWidth) { lineWidth = 0; }
                        lineWidth += width;
                        precedingField = field;
                    }
                    Control flexibleCaption = flexibleField.GetControlFromPosition(0, 0);
                    int minimumWidth = flexibleInput is TextBox ? flexibleInput.MinimumSize.Width : flexibleInput.PreferredSize.Width;
                    if (flexibleCaption != null) { minimumWidth = Math.Max(minimumWidth, flexibleCaption.PreferredSize.Width); }
                    int actionSpace = actionWidth;
                    if (lineWidth + minimumWidth + flexibleField.Margin.Horizontal + actionSpace > availableWidth && precedingField != null)
                    {
                        row.SetFlowBreak(precedingField, true);
                        lineWidth = 0;
                    }
                    bool separateAction = minimumWidth + flexibleField.Margin.Horizontal + actionSpace > availableWidth;
                    row.SetFlowBreak(flexibleField, separateAction);
                    flexibleField.Size = new Size(Math.Max(minimumWidth,
                        availableWidth - lineWidth - (separateAction ? 0 : actionSpace) - flexibleField.Margin.Horizontal), captionHeight + inputHeight);
                    actionField.Size = new Size(separateAction ? availableWidth : actionWidth, captionHeight + inputHeight);
                }
                finally { sizing = false; }
            };
            row.Controls.AddRange(fields);
            return row;
        }

        private static TableLayoutPanel FilterField(string label, Control input)
        {
            var field = new TableLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, ColumnCount = 1, RowCount = 2, Margin = new Padding(0, 0, 12, 0) };
            field.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            field.RowStyles.Add(new RowStyle(SizeType.Absolute));
            field.RowStyles.Add(new RowStyle(SizeType.Absolute));
            if (label != null) { field.Controls.Add(new Label { Text = label, AutoSize = true, Margin = new Padding(0, 0, 0, 4) }, 0, 0); }
            input.Margin = Padding.Empty;
            input.Anchor = input is ComboBox ? AnchorStyles.Top | AnchorStyles.Left : AnchorStyles.Left;
            input.SizeChanged += delegate
            {
                if (!field.IsDisposed && !field.Disposing && field.Parent != null && !field.Parent.IsDisposed && !field.Parent.Disposing)
                { field.Parent.PerformLayout(); }
            };
            field.Controls.Add(input, 0, 1);
            return field;
        }

        private static void ConfigureMetadataFilter(ComboBox filter, string label, IEnumerable<string> values)
        {
            filter.DropDownStyle = ComboBoxStyle.DropDownList;
            if (filter.Parent == null) { filter.Width = 170; }
            filter.AccessibleName = label;
            filter.Items.Clear();
            filter.Items.Add(new GroupingOption(null, Strings.ResourceFilterAny));
            foreach (string value in values.Select(v => (v ?? string.Empty).Trim()).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(v => v, StringComparer.CurrentCultureIgnoreCase))
            { filter.Items.Add(new GroupingOption(value, value.Length == 0 ? Strings.ResourceWithoutInformation : value)); }
            filter.SelectedIndex = 0;
        }

        private static string FilterValue(ComboBox filter)
        { return filter.SelectedItem == null ? null : ((GroupingOption)filter.SelectedItem).Key; }

        private static bool MatchesMetadata(string actual, string filter)
        { return filter == null || string.Equals((actual ?? string.Empty).Trim(), filter, StringComparison.OrdinalIgnoreCase); }

        private void UpdateFloorFilter()
        {
            string previous = FilterValue(_floorFilter);
            _updatingFilters = true;
            try
            {
                ConfigureMetadataFilter(_floorFilter, Strings.ResourceFloorLabel,
                    _resources.Where(r => MatchesMetadata(r.BuildingName, FilterValue(_buildingFilter))).Select(r => r.Floor));
                foreach (GroupingOption option in _floorFilter.Items)
                { if (option.Key == previous) { _floorFilter.SelectedItem = option; break; } }
            }
            finally { _updatingFilters = false; }
        }

        private static IEnumerable<string> FeatureValues(CalendarResource resource)
        {
            // Nextcloud and RoomVox publish room-features as comma-separated tokens.
            return (resource.Features ?? string.Empty).Split(',').Select(value => value.Trim()).Where(value => value.Length != 0);
        }

        private void UpdateFeaturesLabel()
        {
            _featuresFilter.Text = (_selectedFeatures.Count == 0 ? Strings.ResourceFilterAny
                : string.Format(CultureInfo.CurrentCulture, Strings.ResourceFeaturesSelected, _selectedFeatures.Count)) + " ▾";
        }

        private void ClearFilters()
        {
            _updatingFilters = true;
            try
            {
                _search.Clear();
                _buildingFilter.SelectedIndex = 0;
                ConfigureMetadataFilter(_floorFilter, Strings.ResourceFloorLabel, _resources.Select(r => r.Floor));
                _capacityFilter.Value = 0;
                _onlyAvailable.Checked = false;
                foreach (ToolStripMenuItem item in _featuresMenu.Items) { item.Checked = false; }
            }
            finally { _updatingFilters = false; }
            RebuildList();
        }

        private void SetGroupsExpanded(bool expanded)
        {
            _groupsExpanded = expanded;
            _collapsed.Clear();
            _collapsedFloors.Clear();
            _expandAll.Text = expanded ? Strings.ResourceCollapseAll : Strings.ResourceExpandAll;
            RebuildList(false);
        }

        private CalendarResource SelectedResource
        { get { return _list.SelectedItems.Count == 1 ? _list.SelectedItems[0].Tag as CalendarResource : null; } }

        private void UpdateFloorFont()
        {
            Font previous = _floorFont;
            _floorFont = new Font(_list.Font, FontStyle.Bold);
            foreach (ListViewItem item in _list.Items)
            { if (item.Tag is FloorGroup) { item.Font = _floorFont; } }
            if (previous != null) { previous.Dispose(); }
        }

        internal CalendarResourceTimeWindow TimeWindow { get { return _window; } }

        private async Task RefreshTimeWindowAsync()
        {
            if (_closing || IsDisposed || Disposing) { return; }
            CalendarResourceTimeWindow current;
            try { current = _readWindow(); }
            catch (Exception ex)
            {
                if (_closing || IsDisposed || Disposing) { return; }
                DiagnosticsLogger.LogException(LogCategories.Ifb, "Appointment availability window could not be read.", ex);
                _intervalTimer.Stop();
                current = null;
            }
            // Outlook COM calls can process a close/dispose while reading the appointment.
            if (_closing || IsDisposed || Disposing) { return; }
            if (current == null)
            {
                if (_dayView != null) { _dayView.Close(); }
                if (_closing || IsDisposed || Disposing) { return; }
                CancelAvailability();
                _window = null;
                _availability.Clear();
                _interval.Text = Strings.ResourceUnknown;
                RebuildList();
                _apply.Enabled = false;
                return;
            }
            if (current.Matches(_window)) { return; }
            _window = current;
            Task dayUpdate = _dayView != null ? _dayView.UpdateAppointmentWindow(current.StartUtc, current.EndUtc, current.TimeZone) : Task.FromResult(0);
            _interval.Text = Strings.ResourceAppointmentIntervalLabel + ": "
                + TimeZoneInfo.ConvertTimeFromUtc(current.StartUtc, current.TimeZone).ToString("g", CultureInfo.CurrentCulture)
                + " – " + TimeZoneInfo.ConvertTimeFromUtc(current.EndUtc, current.TimeZone).ToString("g", CultureInfo.CurrentCulture)
                + " · " + current.TimeZone.DisplayName;
            await Task.WhenAll(LoadAvailabilityAsync(current), dayUpdate);
        }

        private async Task LoadAvailabilityAsync(CalendarResourceTimeWindow window)
        {
            if (_closing || IsDisposed || Disposing) { return; }
            CancelAvailability();
            var cancellation = new CancellationTokenSource();
            _availabilityCancellation = cancellation;
            int generation = _availabilityGeneration;
            foreach (CalendarResource resource in _resources) { _availability[resource.Email] = CalendarResourceAvailabilityState.Checking; }
            RebuildList();
            try
            {
                var results = await _availabilityService.QueryAsync(_resources, window.StartUtc, window.EndUtc, cancellation.Token);
                if (_closing || IsDisposed || Disposing || cancellation.IsCancellationRequested || generation != _availabilityGeneration) { return; }
                foreach (CalendarResource resource in _resources)
                {
                    CalendarResourceAvailability result;
                    _availability[resource.Email] = results.TryGetValue(resource.Email, out result) ? result.State : CalendarResourceAvailabilityState.Unknown;
                }
            }
            catch (OperationCanceledException) { return; }
            catch (Exception ex)
            {
                if (_closing || IsDisposed || Disposing || generation != _availabilityGeneration) { return; }
                DiagnosticsLogger.LogException(LogCategories.Ifb, "Picker availability query failed.", ex);
                foreach (CalendarResource resource in _resources) { _availability[resource.Email] = CalendarResourceAvailabilityState.Unknown; }
            }
            finally
            {
                if (ReferenceEquals(_availabilityCancellation, cancellation)) { _availabilityCancellation = null; }
                cancellation.Dispose();
            }
            if (!_closing && !IsDisposed && !Disposing && generation == _availabilityGeneration) { RebuildList(); }
        }

        private void CancelAvailability()
        {
            _availabilityGeneration++;
            if (_availabilityCancellation != null) { _availabilityCancellation.Cancel(); }
        }

        private void StopAvailability()
        {
            if (_closing) { return; }
            _closing = true;
            _intervalTimer.Stop();
            CancelAvailability();
        }

        private async Task ShowAvailabilityAsync()
        {
            if (_closing || IsDisposed || Disposing) { return; }
            CalendarResource resource = SelectedResource;
            if (_availabilityService == null || _window == null || resource == null || _openingDayView) { return; }
            _openingDayView = true;
            try
            {
                await RefreshTimeWindowAsync();
                if (_closing || IsDisposed || Disposing || _window == null) { return; }
                using (var dialog = new CalendarResourceAvailabilityForm(resource, _availabilityService, _window.StartUtc, _window.EndUtc, _window.TimeZone))
                {
                    _dayView = dialog;
                    try { dialog.ShowDialog(this); }
                    finally { _dayView = null; }
                }
            }
            finally { _openingDayView = false; }
        }

        private sealed class GroupingOption
        {
            internal readonly string Key;
            private readonly string _label;
            internal GroupingOption(string key, string label) { Key = key; _label = label; }
            public override string ToString() { return _label; }
        }

        private void AdjustInitialSizeForPolicyWarning(Rectangle workingArea)
        {
            if (_initialPolicyWarningSizeApplied) { return; }
            _initialPolicyWarningSizeApplied = true;
            if (!_policyWarningPanel.Visible) { return; }

            int initialHeight = Height;
            int targetHeight = Math.Min(initialHeight + _policyWarningPanel.Height + _policyWarningPanel.Margin.Vertical, workingArea.Height);
            MinimumSize = new Size(MinimumSize.Width, Math.Min(MinimumSize.Height, workingArea.Height));
            Height = targetHeight;
            Top = Math.Max(workingArea.Top, Math.Min(Top - (Height - initialHeight) / 2, workingArea.Bottom - Height));
        }

        private void LayoutPolicyWarningPanel()
        {
            if (_layoutPolicyWarning || IsDisposed || Disposing) { return; }
            _layoutPolicyWarning = true;
            try
            {
                WarningPanelUiHelper.Layout(_policyWarningPanel, _policyWarningTitleLabel,
                    _policyWarningTextLabel, _policyWarningLinkLabel, _policyWarningPanel.Left,
                    _policyWarningPanel.Top, _policyWarningPanel.Width, ScaleLogical(8),
                    ScaleLogical(180), ScaleLogical(4), ScaleLogical(6));
            }
            finally { _layoutPolicyWarning = false; }
        }

        internal List<CalendarResource> SelectedResources
        {
            get { return _resources.Where(r => _selected.Contains(r.Email)).ToList(); }
        }

        private void RebuildList(bool captureGroups = true)
        {
            if (_closing || IsDisposed || Disposing || _updatingFilters) { return; }
            string query = _search.Text.Trim();
            string focusedEmail = SelectedResource == null ? null : SelectedResource.Email;
            var focusedFloor = _list.FocusedItem == null ? null : _list.FocusedItem.Tag as FloorGroup;
            string grouping = ((GroupingOption)_grouping.SelectedItem).Key;
            if (captureGroups) { _list.CaptureGroups(_collapsed); }
            _updating = true;
            _list.BeginUpdate();
            _list.Items.Clear();
            _list.Groups.Clear();
            if (_floorFont == null) { UpdateFloorFont(); }
            UiThemePalette palette = UiThemeManager.DetectPalette();
            var matching = _resources.Where(resource => MatchesFilters(resource, query)).ToList();
            var sections = matching.GroupBy(r => GroupKey(grouping == "building" ? "" : r.Kind, grouping, RawGroupValue(r, grouping)), StringComparer.OrdinalIgnoreCase)
                .OrderBy(g => grouping == "building" ? "" : (g.First().IsRoom ? "0" : "1"))
                .ThenBy(g => RawGroupValue(g.First(), grouping), StringComparer.CurrentCultureIgnoreCase);
            foreach (var section in sections)
            {
                CalendarResource first = section.First();
                string groupValue = RawGroupValue(first, grouping);
                string heading = first.IsRoom ? Strings.ResourceRoomsLabel : Strings.ResourceResourcesLabel;
                if (grouping == "building") { heading = GroupLabel(groupValue); }
                else if (grouping != "none" && (grouping != "room-type" || first.IsRoom)) { heading += " — " + GroupLabel(groupValue); }
                var group = new ListViewGroup(section.Key, heading + " (" + section.Count().ToString(CultureInfo.CurrentCulture) + ")");
                _list.Groups.Add(group);
                if (!_collapsed.ContainsKey(section.Key)) { _collapsed[section.Key] = !_groupsExpanded; }
                if (grouping == "building")
                {
                    foreach (var floor in section.GroupBy(r => (r.Floor ?? string.Empty).Trim(), StringComparer.OrdinalIgnoreCase).OrderBy(g => g.Key, StringComparer.CurrentCultureIgnoreCase))
                    {
                        string key = GroupKey(section.Key, floor.Key);
                        bool collapsed;
                        if (!_collapsedFloors.TryGetValue(key, out collapsed)) { collapsed = !_groupsExpanded; }
                        string action = collapsed ? Strings.ResourceExpandGroup : Strings.ResourceCollapseGroup;
                        string title = (collapsed ? "▶  " : "▼  ") + Strings.ResourceFloorLabel + " " + GroupLabel(floor.Key)
                            + " (" + floor.Count().ToString(CultureInfo.CurrentCulture) + ")";
                        var header = new ListViewItem(new[] { title, "", "", "", "", "" })
                        {
                            Tag = new FloorGroup { Key = key, Collapsed = collapsed }, Group = group, Font = _floorFont,
                            BackColor = SystemInformation.HighContrast ? SystemColors.Window : ControlPaint.Light(_list.BackColor, 0.08f),
                            ForeColor = _list.ForeColor, ToolTipText = action, ImageIndex = -1
                        };
                        _list.Items.Add(header);
                        if (focusedFloor != null && focusedFloor.Key == key) { header.Selected = true; header.Focused = true; }
                        if (!collapsed) { AddResourceRows(floor, group, focusedEmail, palette); }
                    }
                }
                else { AddResourceRows(section, group, focusedEmail, palette); }
            }
            _list.HideFloorCheckboxes();
            _list.RestoreGroups(_collapsed);
            _list.EndUpdate();
            _updating = false;
            UpdateColumnWidths();
            _empty.Text = _onlyAvailable.Checked && _availability.Values.Any(state => state == CalendarResourceAvailabilityState.Checking)
                ? Strings.ResourceChecking : Strings.ResourceNoResults;
            _empty.Visible = matching.Count == 0;
            _apply.Enabled = (_availabilityService == null || _window != null) && SelectedResources.Count(r => r.IsRoom) <= 1;
            if (_list.Items.Count > 0 && _list.SelectedItems.Count == 0) { _list.Items[0].Selected = true; }
            ShowDetails(SelectedResource);
        }

        private bool MatchesFilters(CalendarResource resource, string query)
        {
            string searchable = string.Join(" ", new[] { resource.DisplayName, resource.Email, resource.BuildingName,
                resource.BuildingAddress, resource.Floor, resource.RoomNumber, resource.Features, resource.RoomType, resource.ResourceType });
            if (query.Length != 0 && searchable.IndexOf(query, StringComparison.CurrentCultureIgnoreCase) < 0) { return false; }
            if (!MatchesMetadata(resource.BuildingName, FilterValue(_buildingFilter)) || !MatchesMetadata(resource.Floor, FilterValue(_floorFilter))) { return false; }
            if (_capacityFilter.Value > 0 && (!resource.Capacity.HasValue || resource.Capacity.Value < _capacityFilter.Value)) { return false; }
            if (_selectedFeatures.Count > 0 && !_selectedFeatures.IsSubsetOf(FeatureValues(resource))) { return false; }
            CalendarResourceAvailabilityState state;
            return !_onlyAvailable.Checked || (_availability.TryGetValue(resource.Email, out state) && state == CalendarResourceAvailabilityState.Free);
        }

        private void AddResourceRows(IEnumerable<CalendarResource> resources, ListViewGroup group, string focusedEmail, UiThemePalette palette)
        {
            foreach (CalendarResource resource in resources.OrderBy(r => !r.IsRoom).ThenBy(r => r.DisplayName, StringComparer.CurrentCultureIgnoreCase))
            {
                CalendarResourceAvailabilityState state;
                if (!_availability.TryGetValue(resource.Email, out state)) { state = CalendarResourceAvailabilityState.Unknown; }
                var item = new ListViewItem(ResourceColumnValues(resource))
                { Tag = resource, Checked = _selected.Contains(resource.Email), ImageIndex = resource.IsRoom ? 0 : 1, Group = group,
                    IndentCount = ((GroupingOption)_grouping.SelectedItem).Key == "building" ? 1 : 0, UseItemStyleForSubItems = false };
                ListViewItem.ListViewSubItem status = item.SubItems.Add(AvailabilityText(state));
                item.SubItems.Add("📅");
                foreach (ListViewItem.ListViewSubItem cell in item.SubItems) { cell.BackColor = _list.BackColor; cell.ForeColor = _list.ForeColor; }
                status.ForeColor = SystemInformation.HighContrast ? SystemColors.WindowText : state == CalendarResourceAvailabilityState.Free ? palette.SuccessText
                    : state == CalendarResourceAvailabilityState.Busy ? palette.ErrorText : palette.MutedText;
                _list.Items.Add(item);
                if (resource.Email == focusedEmail) { item.Selected = true; }
            }
        }

        private static string GroupKey(params string[] parts)
        { return string.Join("|", parts.Select(value => value.Length.ToString(CultureInfo.InvariantCulture) + ":" + value)); }

        private static string GroupLabel(string value)
        { return string.IsNullOrWhiteSpace(value) ? Strings.ResourceWithoutInformation : value.Trim(); }

        private static string Metadata(string value) { return string.IsNullOrWhiteSpace(value) ? "—" : value.Trim(); }

        private static string[] ResourceColumnValues(CalendarResource resource)
        {
            return new[]
            {
                string.IsNullOrWhiteSpace(resource.DisplayName) ? resource.Email : resource.DisplayName,
                resource.IsRoom ? Metadata(resource.RoomType) : "—",
                Metadata(resource.BuildingName),
                resource.Capacity.HasValue ? "♟ " + resource.Capacity.Value.ToString(CultureInfo.CurrentCulture) : "—"
            };
        }

        private void MeasureColumnWidths()
        {
            _columnFont = _list.Font;
            _columnDpi = _list.DeviceDpi;
            _minimumColumnWidths = new int[_list.Columns.Count];
            _preferredColumnWidths = new int[_list.Columns.Count];
            int padding = DpiScaling.ScaleLogical(_list, 16);
            using (Graphics graphics = _list.CreateGraphics())
            {
                Func<string, int> measure = text => TextRenderer.MeasureText(graphics, text ?? string.Empty, _list.Font,
                    Size.Empty, TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix).Width + padding;
                for (int index = 0; index < _list.Columns.Count; index++)
                { _minimumColumnWidths[index] = _preferredColumnWidths[index] = measure(_list.Columns[index].Text); }
                int decorations = CheckBoxRenderer.GetGlyphSize(graphics, System.Windows.Forms.VisualStyles.CheckBoxState.UncheckedNormal).Width
                    + (_list.SmallImageList == null ? 0 : _list.SmallImageList.ImageSize.Width * 2) + padding;
                _minimumColumnWidths[0] = Math.Max(_minimumColumnWidths[0] + decorations, DpiScaling.ScaleLogical(_list, 120));
                for (int index = 1; index < 3; index++)
                { _minimumColumnWidths[index] = Math.Max(_minimumColumnWidths[index], DpiScaling.ScaleLogical(_list, 80)); }
                // Measure the complete directory, so filtering and async status changes do not move columns.
                foreach (CalendarResource resource in _resources)
                {
                    string[] values = ResourceColumnValues(resource);
                    for (int index = 0; index < values.Length; index++)
                    { _preferredColumnWidths[index] = Math.Max(_preferredColumnWidths[index], measure(values[index]) + (index == 0 ? decorations : 0)); }
                }
                foreach (CalendarResourceAvailabilityState state in Enum.GetValues(typeof(CalendarResourceAvailabilityState)))
                { _preferredColumnWidths[4] = Math.Max(_preferredColumnWidths[4], measure(AvailabilityText(state))); }
            }
            for (int index = 0; index < _list.Columns.Count; index++)
            { _preferredColumnWidths[index] = Math.Max(_minimumColumnWidths[index], _preferredColumnWidths[index]); }
            _minimumColumnWidths[3] = _preferredColumnWidths[3];
            _minimumColumnWidths[4] = _preferredColumnWidths[4];
        }

        private void UpdateColumnWidths()
        {
            if (_sizingColumns || _updating || _list.IsDisposed || !_list.IsHandleCreated || _list.Columns.Count != 6) { return; }
            _sizingColumns = true;
            try
            {
                if (_preferredColumnWidths == null || _columnFont != _list.Font || _columnDpi != _list.DeviceDpi) { MeasureColumnWidths(); }
                int available = Math.Max(0, _list.ClientSize.Width - DpiScaling.ScaleLogical(_list, 6));
                int[] widths = (int[])_preferredColumnWidths.Clone();
                int preferred = widths.Sum();
                if (available >= preferred) { widths[0] += available - preferred; }
                else
                {
                    widths = (int[])_minimumColumnWidths.Clone();
                    int minimum = widths.Sum();
                    int spare = Math.Max(0, available - minimum);
                    int flexible = preferred - minimum;
                    if (flexible > 0)
                    {
                        for (int index = 0; index < 3; index++)
                        { widths[index] += (int)((long)spare * (_preferredColumnWidths[index] - widths[index]) / flexible); }
                        widths[0] += Math.Max(0, available - widths.Sum());
                    }
                }
                // Below the readable minimum, retain horizontal scrolling instead of crushing headers/status.
                for (int index = 0; index < widths.Length; index++)
                { if (_list.Columns[index].Width != widths[index]) { _list.Columns[index].Width = widths[index]; } }
            }
            finally { _sizingColumns = false; }
        }

        private static string RawGroupValue(CalendarResource resource, string grouping)
        {
            string value;
            switch (grouping)
            {
                case "building": value = resource.BuildingName; break;
                case "floor": value = resource.Floor; break;
                case "room-type": if (!resource.IsRoom) { return string.Empty; } value = resource.RoomType; break;
                default: return string.Empty;
            }
            return (value ?? string.Empty).Trim();
        }

        private static string AvailabilityText(CalendarResourceAvailabilityState state)
        {
            switch (state)
            {
                case CalendarResourceAvailabilityState.Free: return "● " + Strings.ResourceAvailable;
                case CalendarResourceAvailabilityState.Busy: return "● " + Strings.ResourceBusy;
                case CalendarResourceAvailabilityState.Checking: return Strings.ResourceChecking;
                default: return "○ " + Strings.ResourceUnknown;
            }
        }

        private void OnItemCheck(object sender, ItemCheckEventArgs args)
        {
            // Native handle/theme creation replays checkbox notifications; it is not user input.
            if (_updating || _list.ChangingHandle) { return; }
            var resource = _list.Items[args.Index].Tag as CalendarResource;
            if (resource == null) { args.NewValue = args.CurrentValue; return; }
            if (args.NewValue == CheckState.Checked)
            {
                if (resource.IsRoom)
                {
                    foreach (CalendarResource room in _resources.Where(r => r.IsRoom)) { _selected.Remove(room.Email); }
                    _updating = true;
                    foreach (ListViewItem item in _list.Items)
                    {
                        var other = item.Tag as CalendarResource;
                        if (item.Index != args.Index && other != null && other.IsRoom) { item.Checked = false; }
                    }
                    _updating = false;
                }
                _selected.Add(resource.Email);
            }
            else { _selected.Remove(resource.Email); }
            _apply.Enabled = (_availabilityService == null || _window != null) && SelectedResources.Count(r => r.IsRoom) <= 1;
        }

        private void ShowDetails(CalendarResource resource)
        {
            _details.SuspendLayout();
            foreach (Control control in _details.Controls.Cast<Control>().ToArray()) { control.Dispose(); }
            _details.Controls.Clear();
            _details.RowStyles.Clear();
            _details.RowCount = 0;
            if (resource != null)
            {
                AddDetail(Strings.ResourceNameLabel, resource.DisplayName);
                AddDetail(Strings.ResourceKindLabel, resource.IsRoom ? Strings.ResourceRoomKind : Strings.ResourceOtherKind);
                AddDetail(Strings.ResourceCalendarAddressLabel, resource.Email);
                AddDetail(resource.IsRoom ? Strings.ResourceRoomTypeLabel : Strings.ResourceTypeLabel, resource.IsRoom ? resource.RoomType : resource.ResourceType);
                AddDetail(Strings.ResourceCapacityLabel, resource.Capacity.HasValue ? resource.Capacity.Value.ToString(CultureInfo.CurrentCulture) : null);
                AddDetail(Strings.ResourceBuildingLabel, resource.BuildingName);
                AddDetail(Strings.ResourceAddressLabel, resource.BuildingAddress);
                AddDetail(Strings.ResourceFloorLabel, resource.Floor);
                AddDetail(Strings.ResourceRoomNumberLabel, resource.RoomNumber);
                AddDetail(Strings.ResourceFeaturesLabel, resource.Features);
            }
            UiThemeManager.ApplyToControlTree(_details, UiThemeManager.DetectPalette());
            LayoutDetails();
            _details.ResumeLayout();
        }

        private void LayoutDetails()
        {
            if (_layoutDetails || _details.ColumnCount != 2 || _details.Controls.Count == 0) { return; }
            _layoutDetails = true;
            _details.SuspendLayout();
            try
            {
                // Reserve the vertical scrollbar even before it appears, so wrapping cannot change its own width.
                int available = Math.Max(2, _details.Width - _details.Padding.Horizontal - SystemInformation.VerticalScrollBarWidth);
                int labelWidth = _details.Controls.Cast<Control>().Where(c => _details.GetColumn(c) == 0)
                    .Max(c => TextRenderer.MeasureText(c.Text, c.Font, Size.Empty, TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix).Width + c.Margin.Horizontal);
                labelWidth = Math.Min(labelWidth, Math.Max(1, available * 2 / 5));
                _details.ColumnStyles[0].Width = labelWidth;
                foreach (Control control in _details.Controls)
                {
                    int width = Math.Max(1, (_details.GetColumn(control) == 0 ? labelWidth : available - labelWidth) - control.Margin.Horizontal);
                    if (control.MaximumSize.Width != width) { control.MaximumSize = new Size(width, 0); }
                }
            }
            finally { _details.ResumeLayout(); _layoutDetails = false; }
        }

        private void AddDetail(string name, string value)
        {
            int row = _details.RowCount++;
            _details.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _details.Controls.Add(new Label { Text = name, AutoSize = true, Margin = new Padding(0, 4, 8, 8), UseMnemonic = false }, 0, row);
            _details.Controls.Add(new Label { Text = string.IsNullOrWhiteSpace(value) ? Strings.ResourceNotProvided : value, AutoSize = true,
                Margin = new Padding(0, 4, 0, 8), UseMnemonic = false }, 1, row);
        }

        private ImageList CreateResourceIcons(int size)
        {
            var images = new ImageList { ColorDepth = ColorDepth.Depth32Bit, ImageSize = new Size(size, size), TransparentColor = Color.Transparent };
            Color color = SystemInformation.HighContrast ? SystemColors.WindowText : Color.FromArgb(0, 130, 201);
            for (int kind = 0; kind < 2; kind++)
            {
                using (var bitmap = new Bitmap(size, size))
                using (Graphics graphics = Graphics.FromImage(bitmap))
                using (var pen = new Pen(color, 1.7f))
                {
                    graphics.SmoothingMode = SmoothingMode.AntiAlias;
                    graphics.ScaleTransform(size / 24f, size / 24f);
                    pen.StartCap = LineCap.Round;
                    pen.EndCap = LineCap.Round;
                    if (kind == 0)
                    {
                        graphics.DrawRectangle(pen, 5, 3, 13, 18);
                        graphics.DrawLine(pen, 2, 21, 21, 21);
                        graphics.DrawLine(pen, 14, 12, 15, 12);
                    }
                    else
                    {
                        graphics.DrawRectangle(pen, 3, 8, 18, 12);
                        graphics.DrawRectangle(pen, 8, 4, 8, 4);
                        graphics.DrawLine(pen, 3, 13, 21, 13);
                        graphics.DrawLine(pen, 12, 12, 12, 15);
                    }
                    var stored = (Bitmap)bitmap.Clone();
                    _iconImages.Add(stored);
                    images.Images.Add(stored);
                }
            }
            return images;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) { StopAvailability(); }
            if (disposing && _list.SmallImageList != null)
            {
                ImageList images = _list.SmallImageList;
                _list.SmallImageList = null;
                images.Dispose();
            }
            if (disposing)
            {
                _intervalTimer.Dispose();
                _resourceMenu.Dispose();
                _featuresMenu.Dispose();
                _calendarTip.Dispose();
                if (_floorFont != null) { _floorFont.Dispose(); }
                foreach (Bitmap image in _iconImages) { image.Dispose(); }
                _iconImages.Clear();
            }
            base.Dispose(disposing);
        }
    }
}
