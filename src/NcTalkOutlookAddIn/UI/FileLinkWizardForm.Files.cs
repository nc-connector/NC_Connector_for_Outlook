// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    // File-step layout, queue actions, and owner-drawn row rendering.
    internal sealed partial class FileLinkWizardForm
    {
        private void InitializeStepFiles()
        {
            var panel = CreateStepPanel();
            panel.SuspendLayout();

            _fileStepLayout.SuspendLayout();
            _fileStepLayout.ColumnCount = 1;
            _fileStepLayout.RowCount = 3;
            _fileStepLayout.Dock = DockStyle.Fill;
            _fileStepLayout.Padding = new Padding(FileStepPaddingPixels);
            _fileStepLayout.Margin = new Padding(0);
            _fileStepLayout.ColumnStyles.Clear();
            _fileStepLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            _fileStepLayout.RowStyles.Clear();
            _fileStepLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _fileStepLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _fileStepLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            panel.Controls.Add(_fileStepLayout);

            _basePathLabel.Text = Strings.FileLinkWizardBasePathPrefix + (_request.BasePath ?? string.Empty);
            _basePathLabel.AutoSize = true;
            _basePathLabel.Margin = new Padding(0);
            _fileStepLayout.Controls.Add(_basePathLabel, 0, 0);

            _attachmentModeInfoLabel.AutoSize = true;
            _attachmentModeInfoLabel.ForeColor = Color.DimGray;
            _attachmentModeInfoLabel.Visible = false;
            _attachmentModeInfoLabel.Margin = new Padding(0, 8, 0, 0);
            _fileStepLayout.Controls.Add(_attachmentModeInfoLabel, 0, 1);

            _fileStepContentLayout.SuspendLayout();
            _fileStepContentLayout.ColumnCount = 2;
            _fileStepContentLayout.RowCount = 1;
            _fileStepContentLayout.Dock = DockStyle.Fill;
            _fileStepContentLayout.Margin = new Padding(0, 12, 0, 0);
            _fileStepContentLayout.ColumnStyles.Clear();
            _fileStepContentLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
            _fileStepContentLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, FileStepButtonColumnMinWidthPixels));
            _fileStepContentLayout.RowStyles.Clear();
            _fileStepContentLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            _fileStepLayout.Controls.Add(_fileStepContentLayout, 0, 2);

            _fileListView.Dock = DockStyle.Fill;
            _fileListView.Margin = new Padding(0, 0, FileStepButtonColumnSpacingPixels, 0);
            _fileListView.View = View.Details;
            _fileListView.FullRowSelect = true;
            _fileListView.HideSelection = false;
            _fileListView.Scrollable = true;
            _fileListView.OwnerDraw = true;
            _fileListView.Columns.Add(Strings.FileLinkWizardColumnPath, 240);
            _fileListView.Columns.Add(Strings.FileLinkWizardColumnType, 100);
            _fileListView.Columns.Add(Strings.FileLinkWizardColumnStatus, 120);
            _fileListView.Resize += (s, e) => PositionProgressBars();
            _fileListView.DrawColumnHeader += HandleFileListViewDrawColumnHeader;
            _fileListView.DrawItem += HandleFileListViewDrawItem;
            _fileListView.DrawSubItem += HandleFileListViewDrawSubItem;
            _fileListView.HorizontalWheelHandler = HandlePathColumnMouseWheel;
            ConfigureFileListViewRowHeight();
            _fileStepContentLayout.Controls.Add(_fileListView, 0, 0);

            _fileStepActionPanel.FlowDirection = FlowDirection.TopDown;
            _fileStepActionPanel.WrapContents = false;
            _fileStepActionPanel.Dock = DockStyle.Fill;
            _fileStepActionPanel.Margin = new Padding(0);
            _fileStepActionPanel.Padding = new Padding(0);
            _fileStepContentLayout.Controls.Add(_fileStepActionPanel, 1, 0);

            _addFilesButton.Text = Strings.FileLinkWizardAddFilesButton;
            _addFilesButton.AutoSize = false;
            _addFilesButton.Size = new Size(150, 28);
            _addFilesButton.Margin = new Padding(0, 0, 0, FileStepButtonGapPixels);
            _addFilesButton.TextAlign = ContentAlignment.MiddleCenter;
            _addFilesButton.Click += (s, e) => AddFiles();
            _fileStepActionPanel.Controls.Add(_addFilesButton);

            _addFolderButton.Text = Strings.FileLinkWizardAddFolderButton;
            _addFolderButton.AutoSize = false;
            _addFolderButton.Size = new Size(150, 28);
            _addFolderButton.Margin = new Padding(0, 0, 0, FileStepButtonGapPixels);
            _addFolderButton.TextAlign = ContentAlignment.MiddleCenter;
            _addFolderButton.Click += (s, e) => AddFolder();
            _fileStepActionPanel.Controls.Add(_addFolderButton);

            _removeItemButton.Text = Strings.FileLinkWizardRemoveButton;
            _removeItemButton.AutoSize = false;
            _removeItemButton.Size = new Size(150, 28);
            _removeItemButton.Margin = new Padding(0);
            _removeItemButton.TextAlign = ContentAlignment.MiddleCenter;
            _removeItemButton.Click += (s, e) => RemoveSelection();
            _fileStepActionPanel.Controls.Add(_removeItemButton);

            AttachFileQueueDropTarget(panel);
            AttachFileQueueDropTarget(_fileStepLayout);
            AttachFileQueueDropTarget(_fileStepContentLayout);
            AttachFileQueueDropTarget(_fileStepActionPanel);
            AttachFileQueueDropTarget(_fileListView);
            AttachFileQueueDropTarget(_addFilesButton);
            AttachFileQueueDropTarget(_addFolderButton);
            AttachFileQueueDropTarget(_removeItemButton);

            _fileStepContentLayout.ResumeLayout(false);
            _fileStepLayout.ResumeLayout(false);
            _fileStepLayout.PerformLayout();

            panel.ResumeLayout(false);
            panel.PerformLayout();

            panel.ClientSizeChanged += (s, e) => LayoutFileStep(panel.ClientSize);
            LayoutFileStep(panel.ClientSize);
            UpdateQueueColumnWidths();
            PositionProgressBars();

            _steps.Add(panel);
        }

        private void LayoutFileStep(Size clientSize)
        {
            if (_fileStepContentLayout.ColumnStyles.Count >= 2)
            {
                int actionColumnWidth = CalculateFileStepButtonColumnWidth();
                _fileStepContentLayout.ColumnStyles[1].Width = actionColumnWidth;
                ApplyFileStepButtonSize(_addFilesButton, actionColumnWidth);
                ApplyFileStepButtonSize(_addFolderButton, actionColumnWidth);
                ApplyFileStepButtonSize(_removeItemButton, actionColumnWidth);
            }
            int maxInfoWidth = Math.Max(120, clientSize.Width - (FileStepPaddingPixels * 2));
            _attachmentModeInfoLabel.MaximumSize = new Size(maxInfoWidth, 0);

            UpdateQueueColumnWidths();
            PositionProgressBars();
        }

        private void ApplyFileStepButtonSize(Button button, int targetWidth)
        {
            if (button == null)
            {
                return;
            }
            int minWidth;
            FooterButtonLayoutHelper.ApplyButtonSize(button, out minWidth);
            int width = Math.Max(minWidth, Math.Max(ScaleLogical(120), targetWidth));
            button.Size = new Size(width, button.Height);
        }

        private int CalculateFileStepButtonColumnWidth()
        {
            int textPadding = ScaleLogical(40);
            int minWidth = ScaleLogical(FileStepButtonColumnMinWidthPixels);
            int maxTextWidth = Math.Max(
                Math.Max(
                    TextRenderer.MeasureText(_addFilesButton.Text ?? string.Empty, _addFilesButton.Font).Width,
                    TextRenderer.MeasureText(_addFolderButton.Text ?? string.Empty, _addFolderButton.Font).Width),
                TextRenderer.MeasureText(_removeItemButton.Text ?? string.Empty, _removeItemButton.Font).Width);

            return Math.Max(minWidth, maxTextWidth + textPadding);
        }

        private sealed class SelectionUploadState
        {
            internal SelectionUploadState(ListViewItem item)
            {
                Item = item;
                Status = FileLinkUploadStatus.Pending;
            }

            internal ListViewItem Item { get; private set; }

            internal ProgressBar ProgressBar { get; set; }

            internal long TotalBytes { get; set; }

            internal long UploadedBytes { get; set; }

            internal FileLinkUploadStatus Status { get; set; }

            internal string RenamedTo { get; set; }

            internal DateTime UploadStartedUtc { get; set; }

            internal double UploadSpeedKbps { get; set; }
        }

        private sealed class PathScrollableListView : ListView
        {
            private const int WmMouseWheel = 0x020A;

            internal Func<int, bool> HorizontalWheelHandler { get; set; }

            protected override void WndProc(ref Message m)
            {
                if (m.Msg == WmMouseWheel && HorizontalWheelHandler != null)
                {
                    long wParam = m.WParam.ToInt64();
                    int delta = unchecked((short)((wParam >> 16) & 0xffff));
                    if (delta != 0 && HorizontalWheelHandler(delta))
                    {
                        return;
                    }
                }

                base.WndProc(ref m);
            }
        }

        private void AddFiles()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Multiselect = true;
                dialog.CheckFileExists = true;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    var selections = new List<FileLinkSelection>();
                    foreach (string file in dialog.FileNames)
                    {
                        selections.Add(new FileLinkSelection(FileLinkSelectionType.File, file));
                    }

                    AddSelections(selections);
                }
            }
        }

        private void AddFolder()
        {
            using (var dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog(this) == DialogResult.OK && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    AddSelections(new[] { new FileLinkSelection(FileLinkSelectionType.Directory, dialog.SelectedPath) });
                }
            }
        }

        private void RemoveSelection()
        {
            if (_fileListView.SelectedItems.Count == 0)
            {
                return;
            }
            foreach (ListViewItem item in _fileListView.SelectedItems)
            {
                FileLinkSelection selection = item.Tag as FileLinkSelection;
                if (selection != null)
                {
                    _items.Remove(selection);
                    SelectionUploadState state;
                    if (_selectionStates.TryGetValue(selection, out state))
                    {
                        DisposeStateProgressBar(state);
                        _selectionStates.Remove(selection);
                    }
                }
                _fileListView.Items.Remove(item);
            }
            if (_items.Count == 0)
            {
                _allowEmptyUpload = false;
            }
            UpdateQueueColumnWidths();
            PositionProgressBars();
            InvalidateUpload();
        }

        private void AddSelections(IEnumerable<FileLinkSelection> selections)
        {
            if (selections == null)
            {
                return;
            }
            var pendingSelections = selections.Where(s => s != null && !string.IsNullOrWhiteSpace(s.LocalPath)).ToList();
            if (pendingSelections.Count == 0)
            {
                return;
            }
            var existingPaths = _attachmentMode
                ? null
                : new HashSet<string>(
                    _items.Select(i => i.LocalPath ?? string.Empty),
                    StringComparer.OrdinalIgnoreCase);

            int requestedCount = pendingSelections.Count;
            int addedCount = 0;

            _fileListView.BeginUpdate();
            try
            {
                foreach (var selection in pendingSelections)
                {
                    if (TryAddSelection(selection, existingPaths))
                    {
                        addedCount++;
                    }
                }
            }
            finally
            {
                _fileListView.EndUpdate();
            }
            if (addedCount == 0)
            {
                return;
            }

            _allowEmptyUpload = false;
            DiagnosticsLogger.Log(
                LogCategories.FileLink,
                "Queue selections added (requested="
                + requestedCount.ToString(CultureInfo.InvariantCulture)
                + ", added="
                + addedCount.ToString(CultureInfo.InvariantCulture)
                + ", total="
                + _items.Count.ToString(CultureInfo.InvariantCulture)
                + ").");

            UpdateQueueColumnWidths();
            PositionProgressBars();
            InvalidateUpload();
        }

        private void UpdateQueueColumnWidths()
        {
            if (_fileListView == null || _fileListView.Columns.Count < 3 || _fileListView.IsDisposed || _fileListView.Disposing)
            {
                return;
            }
            int typeWidth = 110;
            int statusWidth = 180;
            int clientWidth = Math.Max(0, _fileListView.ClientSize.Width);
            int pathWidth = clientWidth - typeWidth - statusWidth - 6;
            if (pathWidth < 120)
            {
                int shortage = 120 - pathWidth;
                int reducibleStatus = Math.Max(0, statusWidth - 150);
                int reduceStatus = Math.Min(shortage, reducibleStatus);
                statusWidth -= reduceStatus;
                shortage -= reduceStatus;

                int reducibleType = Math.Max(0, typeWidth - 90);
                int reduceType = Math.Min(shortage, reducibleType);
                typeWidth -= reduceType;

                pathWidth = Math.Max(90, clientWidth - typeWidth - statusWidth - 6);
            }

            _fileListView.Columns[0].Width = pathWidth;
            _fileListView.Columns[1].Width = typeWidth;
            _fileListView.Columns[2].Width = statusWidth;
            UpdatePathColumnScrollRange();
            _fileListView.Invalidate();
        }

        private bool HandlePathColumnMouseWheel(int delta)
        {
            if (delta == 0)
            {
                return false;
            }
            if (_pathColumnMaxHorizontalOffset <= 0)
            {
                _pathColumnHorizontalOffset = 0;
                return false;
            }
            int steps = Math.Max(1, Math.Abs(delta) / 120);
            int shift = steps * PathColumnWheelStepPixels;
            int nextOffset = _pathColumnHorizontalOffset + (delta < 0 ? shift : -shift);
            if (nextOffset < 0)
            {
                nextOffset = 0;
            }
            else if (nextOffset > _pathColumnMaxHorizontalOffset)
            {
                nextOffset = _pathColumnMaxHorizontalOffset;
            }
            if (nextOffset == _pathColumnHorizontalOffset)
            {
                return true;
            }

            _pathColumnHorizontalOffset = nextOffset;
            _fileListView.Invalidate();
            return true;
        }

        private void UpdatePathColumnScrollRange()
        {
            if (_fileListView == null || _fileListView.IsDisposed || _fileListView.Disposing || _fileListView.Columns.Count == 0)
            {
                _pathColumnHorizontalOffset = 0;
                _pathColumnMaxHorizontalOffset = 0;
                return;
            }
            int visibleWidth = Math.Max(0, _fileListView.Columns[0].Width - 8);
            if (_fileListView.Items.Count == 0 || visibleWidth <= 0)
            {
                _pathColumnHorizontalOffset = 0;
                _pathColumnMaxHorizontalOffset = 0;
                return;
            }
            int widestPath = 0;
            foreach (ListViewItem item in _fileListView.Items)
            {
                string path = item != null ? (item.Text ?? string.Empty) : string.Empty;
                if (path.Length == 0)
                {
                    continue;
                }
                int width = TextRenderer.MeasureText(
                    path,
                    _fileListView.Font,
                    new Size(int.MaxValue, int.MaxValue),
                    TextFormatFlags.SingleLine | TextFormatFlags.NoPadding).Width;

                if (width > widestPath)
                {
                    widestPath = width;
                }
            }

            _pathColumnMaxHorizontalOffset = Math.Max(0, widestPath - visibleWidth);
            if (_pathColumnHorizontalOffset > _pathColumnMaxHorizontalOffset)
            {
                _pathColumnHorizontalOffset = _pathColumnMaxHorizontalOffset;
            }
        }

        private void HandleFileListViewDrawColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e)
        {
            if (e == null)
            {
                return;
            }

            e.DrawDefault = true;
        }

        private void HandleFileListViewDrawItem(object sender, DrawListViewItemEventArgs e)
        {
            if (e == null)
            {
                return;
            }
            if (_fileListView.View != View.Details)
            {
                e.DrawDefault = true;
            }
        }

        private void HandleFileListViewDrawSubItem(object sender, DrawListViewSubItemEventArgs e)
        {
            if (e == null || e.Item == null || e.SubItem == null)
            {
                return;
            }

            Color rowBackColor = e.Item.BackColor.IsEmpty ? _fileListView.BackColor : e.Item.BackColor;
            Color rowTextColor = e.Item.ForeColor.IsEmpty ? _fileListView.ForeColor : e.Item.ForeColor;
            Color backColor = e.SubItem.BackColor.IsEmpty ? rowBackColor : e.SubItem.BackColor;
            Color textColor = e.SubItem.ForeColor.IsEmpty ? rowTextColor : e.SubItem.ForeColor;

            bool selected = e.Item.Selected && (!_fileListView.HideSelection || _fileListView.Focused);
            if (selected)
            {
                backColor = _themePalette != null ? _themePalette.SelectionBackground : SystemColors.Highlight;
                textColor = _themePalette != null ? _themePalette.SelectionText : SystemColors.HighlightText;
            }

            using (var backBrush = new SolidBrush(backColor))
            {
                e.Graphics.FillRectangle(backBrush, e.Bounds);
            }
            string text = e.SubItem.Text ?? string.Empty;
            Rectangle textBounds = new Rectangle(
                e.Bounds.Left + 4,
                e.Bounds.Top + 1,
                Math.Max(0, e.Bounds.Width - 6),
                Math.Max(0, e.Bounds.Height - 2));

            if (e.ColumnIndex == 2)
            {
                SelectionUploadState selectionState = ResolveSelectionState(e.Item);
                if (selectionState != null && selectionState.Status == FileLinkUploadStatus.Uploading)
                {
                    int topPadding = ScaleLogical(2);
                    int bottomPadding = ScaleLogical(2);
                    int barHeight = Math.Max(ScaleLogical(6), 6);
                    int contentHeight = Math.Max(0, e.Bounds.Height - topPadding - bottomPadding);
                    barHeight = Math.Min(barHeight, contentHeight);
                    int speedTop = e.Bounds.Top + topPadding + barHeight + 1;
                    int speedBottom = e.Bounds.Bottom - bottomPadding;
                    int speedHeight = Math.Max(0, speedBottom - speedTop);
                    if (speedHeight > 0)
                    {
                        textBounds = new Rectangle(
                            e.Bounds.Left + 4,
                            speedTop,
                            Math.Max(0, e.Bounds.Width - 6),
                            speedHeight);
                    }
                }
            }

            if (e.ColumnIndex == 0)
            {
                var state = e.Graphics.Save();
                e.Graphics.SetClip(e.Bounds);
                var shiftedTextBounds = new Rectangle(
                    textBounds.Left - _pathColumnHorizontalOffset,
                    textBounds.Top,
                    textBounds.Width + _pathColumnHorizontalOffset,
                    textBounds.Height);
                TextRenderer.DrawText(
                    e.Graphics,
                    text,
                    _fileListView.Font,
                    shiftedTextBounds,
                    textColor,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPadding);
                e.Graphics.Restore(state);
            }
            else
            {
                TextRenderer.DrawText(
                    e.Graphics,
                    text,
                    _fileListView.Font,
                    textBounds,
                    textColor,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPadding);
            }
            if (e.ColumnIndex == _fileListView.Columns.Count - 1 && selected)
            {
                Rectangle focusRect = e.Item.Bounds;
                focusRect.Width = Math.Max(0, _fileListView.ClientSize.Width - focusRect.Left);
                ControlPaint.DrawFocusRectangle(e.Graphics, focusRect, textColor, backColor);
            }
        }

        private SelectionUploadState ResolveSelectionState(ListViewItem item)
        {
            if (item == null)
            {
                return null;
            }

            var selection = item.Tag as FileLinkSelection;
            if (selection == null)
            {
                return null;
            }

            SelectionUploadState selectionState;
            if (_selectionStates.TryGetValue(selection, out selectionState))
            {
                return selectionState;
            }

            return null;
        }

        private void ApplyQueueRowStyle(SelectionUploadState state, Color backgroundColor, Color textColor)
        {
            if (state == null || state.Item == null)
            {
                return;
            }

            ListViewItem item = state.Item;
            item.BackColor = backgroundColor;
            item.ForeColor = textColor;

            for (int i = 0; i < item.SubItems.Count; i++)
            {
                item.SubItems[i].BackColor = backgroundColor;
                if (i != 2)
                {
                    item.SubItems[i].ForeColor = textColor;
                }
            }
        }

        private void ConfigureFileListViewRowHeight()
        {
            if (_fileListView == null || _fileListView.IsDisposed || _fileListView.Disposing)
            {
                return;
            }
            int rowHeight = Math.Max(ScaleLogical(30), 30);
            _fileListRowHeightImageList.ColorDepth = ColorDepth.Depth32Bit;
            _fileListRowHeightImageList.ImageSize = new Size(1, rowHeight);
            _fileListRowHeightImageList.Images.Clear();
            _fileListRowHeightImageList.Images.Add(new Bitmap(1, rowHeight));
            _fileListView.SmallImageList = _fileListRowHeightImageList;
        }
    }
}
