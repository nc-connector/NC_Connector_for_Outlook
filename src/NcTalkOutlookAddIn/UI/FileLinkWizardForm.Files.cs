// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    // File-step layout and source selection.
    internal sealed partial class FileLinkWizardForm
    {
        private void InitializeStepFiles()
        {
            var panel = CreateStepPanel();
            panel.SuspendLayout();

            _fileStepLayout.SuspendLayout();
            _fileStepLayout.ColumnCount = 1;
            _fileStepLayout.RowCount = 5;
            _fileStepLayout.Dock = DockStyle.Fill;
            _fileStepLayout.Padding = new Padding(FileStepPaddingPixels);
            _fileStepLayout.Margin = new Padding(0);
            _fileStepLayout.ColumnStyles.Clear();
            _fileStepLayout.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100f));
            _fileStepLayout.RowStyles.Clear();
            _fileStepLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _fileStepLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _fileStepLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            _fileStepLayout.RowStyles.Add(
                new RowStyle(SizeType.Absolute, ScaleLogical(32)));
            _fileStepLayout.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100f));
            panel.Controls.Add(_fileStepLayout);

            var destinationPanel = new TableLayoutPanel
            {
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Top,
                ColumnCount = 1,
                RowCount = 2,
                Margin = new Padding(0)
            };
            destinationPanel.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100f));
            destinationPanel.RowStyles.Add(
                new RowStyle(SizeType.AutoSize));
            destinationPanel.RowStyles.Add(
                new RowStyle(SizeType.AutoSize));

            _targetFolderCaptionLabel.Text =
                Strings.FileLinkQueueTargetFolder;
            _targetFolderCaptionLabel.AutoSize = true;
            _targetFolderCaptionLabel.Margin = new Padding(0);
            destinationPanel.Controls.Add(
                _targetFolderCaptionLabel,
                0,
                0);

            _basePathLabel.AutoSize = true;
            _basePathLabel.Font = new Font(
                _basePathLabel.Font,
                FontStyle.Bold);
            _basePathLabel.Margin = new Padding(0, ScaleLogical(3), 0, 0);
            destinationPanel.Controls.Add(_basePathLabel, 0, 1);
            _fileStepLayout.Controls.Add(destinationPanel, 0, 0);

            _attachmentModeInfoLabel.AutoSize = true;
            _attachmentModeInfoLabel.ForeColor = _themePalette.MutedText;
            _attachmentModeInfoLabel.Visible = false;
            _attachmentModeInfoLabel.Margin = new Padding(
                0,
                ScaleLogical(7),
                0,
                0);
            _fileStepLayout.Controls.Add(
                _attachmentModeInfoLabel,
                0,
                1);

            InitializeQueueSourceActions();
            _fileStepLayout.Controls.Add(_fileStepActionPanel, 0, 2);

            InitializeQueueSummary();
            _fileStepLayout.Controls.Add(_queueSummaryPanel, 0, 3);

            InitializeQueueList();
            _fileStepLayout.Controls.Add(_fileStepContentLayout, 0, 4);

            AttachFileQueueDropTarget(panel);
            AttachFileQueueDropTarget(_fileStepLayout);
            AttachFileQueueDropTarget(_fileStepContentLayout);
            AttachFileQueueDropTarget(_fileStepActionPanel);
            AttachFileQueueDropTarget(_fileListView);
            AttachFileQueueDropTarget(_queueEmptyLabel);
            AttachFileQueueDropTarget(_localSourceButton);
            AttachFileQueueDropTarget(_nextcloudSourceButton);

            _fileStepLayout.ResumeLayout(false);
            _fileStepLayout.PerformLayout();
            panel.ResumeLayout(false);
            panel.PerformLayout();

            panel.ClientSizeChanged +=
                (s, e) => LayoutFileStep(panel.ClientSize);
            RefreshQueueTargetPath();
            RebuildQueueView();
            LayoutFileStep(panel.ClientSize);

            _steps.Add(panel);
        }

        private void InitializeQueueSourceActions()
        {
            _fileStepActionPanel.FlowDirection = FlowDirection.LeftToRight;
            _fileStepActionPanel.WrapContents = false;
            _fileStepActionPanel.AutoSize = true;
            _fileStepActionPanel.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            _fileStepActionPanel.Dock = DockStyle.Top;
            _fileStepActionPanel.Margin = new Padding(
                0,
                ScaleLogical(12),
                0,
                ScaleLogical(8));
            _fileStepActionPanel.Padding = new Padding(0);

            ConfigureSourceButton(
                _localSourceButton,
                Strings.FileLinkQueueAddLocal,
                FileLinkIconProvider.LocalSourceKey,
                _localSourceMenu);
            _fileStepActionPanel.Controls.Add(_localSourceButton);

            ConfigureSourceButton(
                _nextcloudSourceButton,
                Strings.FileLinkQueueAddNextcloud,
                FileLinkIconProvider.NextcloudSourceKey,
                _nextcloudSourceMenu);
            _fileStepActionPanel.Controls.Add(_nextcloudSourceButton);

            AddSourceMenuItems(
                _localSourceMenu,
                AddFiles,
                AddFolder);
            AddSourceMenuItems(
                _nextcloudSourceMenu,
                AddNextcloudFiles,
                AddNextcloudFolder);
        }

        private void ConfigureSourceButton(
            Button button,
            string text,
            string imageKey,
            ContextMenuStrip menu)
        {
            button.Text = text;
            button.AutoSize = false;
            button.Size = new Size(
                ScaleLogical(190),
                ScaleLogical(36));
            button.Margin = new Padding(
                0,
                0,
                ScaleLogical(FileStepButtonGapPixels),
                0);
            button.TextAlign = ContentAlignment.MiddleLeft;
            button.ImageAlign = ContentAlignment.MiddleLeft;
            button.TextImageRelation = TextImageRelation.ImageBeforeText;
            button.Image = CreateSourceButtonImage(
                _fileQueueImageList.Images[imageKey],
                ScaleLogical(6));
            button.Click += (s, e) => menu.Show(
                button,
                new Point(0, button.Height));
            button.Paint += DrawSourceButtonArrow;
        }

        private static Bitmap CreateSourceButtonImage(
            Image source,
            int trailingSpace)
        {
            if (source == null)
            {
                return null;
            }
            var image = new Bitmap(
                source.Width + Math.Max(0, trailingSpace),
                source.Height);
            using (Graphics graphics = Graphics.FromImage(image))
            {
                graphics.DrawImageUnscaled(source, 0, 0);
            }
            return image;
        }

        private void AddSourceMenuItems(
            ContextMenuStrip menu,
            Func<Task> addFiles,
            Func<Task> addFolder)
        {
            menu.BackColor = _themePalette.ControlBackground;
            menu.ForeColor = _themePalette.Text;
            menu.Renderer = new ToolStripProfessionalRenderer(
                new QueueMenuColorTable(_themePalette));
            var filesItem = new ToolStripMenuItem(
                Strings.FileLinkQueueSelectFiles,
                new Bitmap(
                    _fileQueueImageList.Images[
                        FileLinkIconProvider.GenericFileKey]),
                async (s, e) => await addFiles());
            var folderItem = new ToolStripMenuItem(
                Strings.FileLinkQueueSelectFolder,
                new Bitmap(
                    _fileQueueImageList.Images[
                        FileLinkIconProvider.FolderKey]),
                async (s, e) => await addFolder());
            menu.Items.Add(filesItem);
            menu.Items.Add(folderItem);
        }

        private void ApplyQueueTheme()
        {
            _queueSummaryPanel.BackColor =
                _themePalette.ControlBackground;
            _queueSummaryPanel.ForeColor = _themePalette.Text;
            _queueSummaryLabel.BackColor =
                _themePalette.ControlBackground;
            _queueSummaryLabel.ForeColor = _themePalette.MutedText;
            _queueStorageLabel.BackColor =
                _themePalette.ControlBackground;
            _queueStorageLabel.ForeColor = _themePalette.MutedText;
            _queueEmptyLabel.BackColor = _themePalette.InputBackground;
            _queueEmptyLabel.ForeColor = _themePalette.MutedText;
            _fileListView.BackColor = _themePalette.InputBackground;
            _fileListView.ForeColor = _themePalette.Text;
            foreach (ContextMenuStrip menu in new[]
            {
                _localSourceMenu,
                _nextcloudSourceMenu
            })
            {
                menu.BackColor = _themePalette.ControlBackground;
                menu.ForeColor = _themePalette.Text;
                foreach (ToolStripItem item in menu.Items)
                {
                    item.BackColor = _themePalette.ControlBackground;
                    item.ForeColor = _themePalette.Text;
                }
            }
        }

        private sealed class QueueMenuColorTable
            : ProfessionalColorTable
        {
            private readonly UiThemePalette _palette;

            internal QueueMenuColorTable(UiThemePalette palette)
            {
                _palette = palette;
                UseSystemColors = false;
            }

            public override Color ToolStripDropDownBackground
            {
                get { return _palette.ControlBackground; }
            }

            public override Color MenuItemSelected
            {
                get { return _palette.SelectionBackground; }
            }

            public override Color MenuItemBorder
            {
                get { return _palette.Border; }
            }

            public override Color ImageMarginGradientBegin
            {
                get { return _palette.ControlBackground; }
            }

            public override Color ImageMarginGradientMiddle
            {
                get { return _palette.ControlBackground; }
            }

            public override Color ImageMarginGradientEnd
            {
                get { return _palette.ControlBackground; }
            }

            public override Color SeparatorDark
            {
                get { return _palette.Border; }
            }

            public override Color SeparatorLight
            {
                get { return _palette.Border; }
            }
        }

        private void DrawSourceButtonArrow(object sender, PaintEventArgs e)
        {
            var button = sender as Button;
            if (button == null || e == null)
            {
                return;
            }
            string arrow = "▾";
            Size size = TextRenderer.MeasureText(
                arrow,
                button.Font,
                Size.Empty,
                TextFormatFlags.NoPadding);
            var bounds = new Rectangle(
                Math.Max(0, button.ClientSize.Width - size.Width - ScaleLogical(9)),
                Math.Max(0, (button.ClientSize.Height - size.Height) / 2),
                size.Width,
                size.Height);
            TextRenderer.DrawText(
                e.Graphics,
                arrow,
                button.Font,
                bounds,
                button.Enabled
                    ? button.ForeColor
                    : _themePalette.DisabledText,
                TextFormatFlags.NoPadding);
        }

        private void InitializeQueueSummary()
        {
            _queueSummaryPanel.ColumnCount = 2;
            _queueSummaryPanel.RowCount = 1;
            _queueSummaryPanel.Dock = DockStyle.Fill;
            _queueSummaryPanel.Margin = new Padding(0);
            _queueSummaryPanel.Padding = new Padding(
                ScaleLogical(8),
                ScaleLogical(4),
                ScaleLogical(8),
                ScaleLogical(4));
            _queueSummaryPanel.BorderStyle = BorderStyle.FixedSingle;
            _queueSummaryPanel.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 55f));
            _queueSummaryPanel.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 45f));
            _queueSummaryPanel.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100f));

            _queueSummaryLabel.Dock = DockStyle.Fill;
            _queueSummaryLabel.AutoEllipsis = true;
            _queueSummaryLabel.TextAlign = ContentAlignment.MiddleLeft;
            _queueSummaryLabel.ForeColor = _themePalette.MutedText;
            _queueSummaryLabel.Margin = new Padding(0);
            _queueSummaryPanel.Controls.Add(_queueSummaryLabel, 0, 0);

            _queueStorageLabel.Dock = DockStyle.Fill;
            _queueStorageLabel.AutoEllipsis = true;
            _queueStorageLabel.TextAlign = ContentAlignment.MiddleRight;
            _queueStorageLabel.ForeColor = _themePalette.MutedText;
            _queueStorageLabel.Margin = new Padding(0);
            _queueSummaryPanel.Controls.Add(_queueStorageLabel, 1, 0);
        }

        private void InitializeQueueList()
        {
            _fileStepContentLayout.SuspendLayout();
            _fileStepContentLayout.ColumnCount = 1;
            _fileStepContentLayout.RowCount = 1;
            _fileStepContentLayout.Dock = DockStyle.Fill;
            _fileStepContentLayout.Margin = new Padding(
                0,
                ScaleLogical(8),
                0,
                0);
            _fileStepContentLayout.ColumnStyles.Clear();
            _fileStepContentLayout.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100f));
            _fileStepContentLayout.RowStyles.Clear();
            _fileStepContentLayout.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100f));

            _fileListView.Dock = DockStyle.Fill;
            _fileListView.Margin = new Padding(0);
            _fileListView.View = View.Details;
            _fileListView.HeaderStyle = ColumnHeaderStyle.None;
            _fileListView.FullRowSelect = true;
            _fileListView.HideSelection = false;
            _fileListView.MultiSelect = true;
            _fileListView.Scrollable = true;
            _fileListView.OwnerDraw = true;
            _fileListView.SmallImageList = _fileQueueImageList;
            _fileListView.ShowItemToolTips = true;
            _fileListView.Columns.Add(string.Empty, ScaleLogical(340));
            _fileListView.Columns.Add(string.Empty, ScaleLogical(90));
            _fileListView.Columns.Add(string.Empty, ScaleLogical(132));
            _fileListView.Columns.Add(string.Empty, ScaleLogical(36));
            _fileListView.Resize += (s, e) =>
            {
                UpdateQueueColumnWidths();
                PositionProgressBars();
            };
            _fileListView.DrawColumnHeader +=
                HandleFileListViewDrawColumnHeader;
            _fileListView.DrawItem += HandleFileListViewDrawItem;
            _fileListView.DrawSubItem += HandleFileListViewDrawSubItem;
            _fileListView.MouseClick += HandleQueueMouseClick;
            _fileListView.MouseDoubleClick += HandleQueueMouseDoubleClick;
            _fileListView.MouseMove += HandleQueueMouseMove;
            _fileListView.KeyDown += HandleQueueKeyDown;
            _fileStepContentLayout.Controls.Add(_fileListView, 0, 0);

            _queueEmptyLabel.Text = Strings.FileLinkQueueEmpty;
            _queueEmptyLabel.Dock = DockStyle.Fill;
            _queueEmptyLabel.TextAlign = ContentAlignment.MiddleCenter;
            _queueEmptyLabel.ForeColor = _themePalette.MutedText;
            _queueEmptyLabel.BackColor = _themePalette.InputBackground;
            _queueEmptyLabel.BorderStyle = BorderStyle.FixedSingle;
            _queueEmptyLabel.Margin = new Padding(0);
            _fileStepContentLayout.Controls.Add(_queueEmptyLabel, 0, 0);

            _fileStepContentLayout.ResumeLayout(false);
        }

        private void LayoutFileStep(Size clientSize)
        {
            int textPadding = ScaleLogical(52);
            int minimumButtonWidth = ScaleLogical(170);
            int widestText = Math.Max(
                TextRenderer.MeasureText(
                    _localSourceButton.Text,
                    _localSourceButton.Font).Width,
                TextRenderer.MeasureText(
                    _nextcloudSourceButton.Text,
                    _nextcloudSourceButton.Font).Width);
            int buttonWidth = Math.Max(
                minimumButtonWidth,
                widestText + textPadding);
            _localSourceButton.Width = buttonWidth;
            _nextcloudSourceButton.Width = buttonWidth;

            int maxInfoWidth = Math.Max(
                ScaleLogical(120),
                clientSize.Width - ScaleLogical(FileStepPaddingPixels * 2));
            _attachmentModeInfoLabel.MaximumSize = new Size(
                maxInfoWidth,
                0);
            _basePathLabel.MaximumSize = new Size(maxInfoWidth, 0);

            UpdateQueueColumnWidths();
            PositionProgressBars();
        }

        private sealed class SelectionUploadState
        {
            internal SelectionUploadState(FileLinkSelection selection)
            {
                Selection = selection;
                Status = FileLinkUploadStatus.Pending;
            }

            internal FileLinkSelection Selection { get; private set; }

            internal ListViewItem Item { get; set; }

            internal ProgressBar ProgressBar { get; set; }

            internal long TotalBytes { get; set; }

            internal long UploadedBytes { get; set; }

            internal FileLinkUploadStatus Status { get; set; }

            internal string RenamedTo { get; set; }

            internal DateTime UploadStartedUtc { get; set; }

            internal double UploadSpeedKbps { get; set; }
        }

        private enum FileLinkQueueRowKind
        {
            Source,
            Folder,
            File
        }

        private sealed class FileLinkQueueRow
        {
            internal FileLinkQueueRow(
                FileLinkQueueRowKind kind,
                FileLinkSelectionSource source,
                FileLinkSelection selection,
                FileLinkQueueNode node,
                int depth,
                string key,
                bool expanded,
                bool canRemove)
            {
                Kind = kind;
                Source = source;
                Selection = selection;
                Node = node;
                Depth = depth;
                Key = key ?? string.Empty;
                Expanded = expanded;
                CanRemove = canRemove;
            }

            internal FileLinkQueueRowKind Kind { get; private set; }

            internal FileLinkSelectionSource Source { get; private set; }

            internal FileLinkSelection Selection { get; private set; }

            internal FileLinkQueueNode Node { get; private set; }

            internal int Depth { get; private set; }

            internal string Key { get; private set; }

            internal bool Expanded { get; private set; }

            internal bool CanRemove { get; private set; }

            internal bool HasChildren
            {
                get
                {
                    return Node != null
                           && Node.IsDirectory
                           && Node.Children.Count > 0;
                }
            }
        }

        private sealed class PathScrollableListView : ListView
        {
        }

        private async Task AddFiles()
        {
            using (var dialog = new OpenFileDialog())
            {
                dialog.Multiselect = true;
                dialog.CheckFileExists = true;
                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    await AddSelectionsAsync(
                        dialog.FileNames.Select(
                            file => new FileLinkSelection(
                                FileLinkSelectionType.File,
                                file)));
                }
            }
        }

        private async Task AddFolder()
        {
            using (var dialog = new FolderBrowserDialog())
            {
                if (dialog.ShowDialog(this) == DialogResult.OK
                    && !string.IsNullOrWhiteSpace(dialog.SelectedPath))
                {
                    await AddSelectionsAsync(
                        new[]
                        {
                            new FileLinkSelection(
                                FileLinkSelectionType.Directory,
                                dialog.SelectedPath)
                        });
                }
            }
        }

        private async Task AddNextcloudFiles()
        {
            using (var dialog = new NextcloudFilePickerForm(
                _service,
                NextcloudFilePickerMode.Files))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }
                await AddSelectionsAsync(
                    dialog.SelectedEntries
                        .Where(entry => entry != null && !entry.IsDirectory)
                        .Select(FileLinkSelection.FromNextcloudFile));
            }
        }

        private async Task AddNextcloudFolder()
        {
            using (var dialog = new NextcloudFilePickerForm(
                _service,
                NextcloudFilePickerMode.Folder))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK
                    || dialog.SelectedEntries.Count == 0)
                {
                    return;
                }
                NextcloudStorageEntry root = dialog.SelectedEntries[0];
                await AddSelectionsAsync(
                    new[]
                    {
                        FileLinkSelection.FromNextcloudFolder(
                            root,
                            dialog.SelectedEntries.Skip(1))
                    });
            }
        }

        private void RemoveSelection()
        {
            var selections = _fileListView.SelectedItems
                .Cast<ListViewItem>()
                .Select(item => item.Tag as FileLinkQueueRow)
                .Where(row => row != null && row.CanRemove)
                .Select(row => row.Selection)
                .Where(selection => selection != null)
                .Distinct()
                .ToList();
            RemoveSelections(selections);
        }

        private void RemoveSelections(
            IEnumerable<FileLinkSelection> selections)
        {
            if (IsWizardBusy || selections == null)
            {
                return;
            }
            List<FileLinkSelection> removed = selections
                .Where(selection => selection != null)
                .Distinct()
                .ToList();
            if (removed.Count == 0)
            {
                return;
            }

            foreach (FileLinkSelection selection in removed)
            {
                _items.Remove(selection);
                _queueSnapshots.Remove(selection);
                SelectionUploadState state;
                if (_selectionStates.TryGetValue(selection, out state))
                {
                    DisposeStateProgressBar(state);
                    _selectionStates.Remove(selection);
                }
                string keyPrefix = selection.IdentityPath + "|";
                _expandedQueueFolders.RemoveWhere(
                    key => key.StartsWith(
                        keyPrefix,
                        StringComparison.Ordinal));
            }
            if (_items.Count == 0)
            {
                _allowEmptyUpload = false;
            }
            RebuildQueueView();
            InvalidateUpload();
        }

        private async Task AddSelectionsAsync(
            IEnumerable<FileLinkSelection> selections)
        {
            if (IsWizardBusy || selections == null)
            {
                return;
            }
            var pendingSelections = selections
                .Where(
                    selection => selection != null
                                 && (selection.Source
                                     == FileLinkSelectionSource.Nextcloud
                                     || !string.IsNullOrWhiteSpace(
                                         selection.LocalPath)))
                .ToList();
            if (pendingSelections.Count == 0)
            {
                return;
            }
            var existingSelections = _attachmentMode
                ? null
                : new HashSet<FileLinkSelection>(
                    _items,
                    FileLinkSelection.IdentityComparer);

            int requestedCount = pendingSelections.Count;
            int addedCount = 0;
            Exception firstFailure = null;
            bool scanCancelled = false;
            _queueScanInProgress = true;
            _progressBar.Style = ProgressBarStyle.Marquee;
            _progressLabel.Text = Strings.FileLinkWizardStatusScanning;
            SetProgressPanelVisible(true);
            UpdateNavigationState();
            UpdateUploadButtonState();
            _cancellationSource = new CancellationTokenSource();
            CancellationToken token = _cancellationSource.Token;
            try
            {
                foreach (FileLinkSelection selection in pendingSelections)
                {
                    token.ThrowIfCancellationRequested();
                    if (!TryReserveSelection(
                        selection,
                        existingSelections))
                    {
                        continue;
                    }
                    try
                    {
                        FileLinkQueueNode snapshot = await Task.Run(
                            () => FileLinkQueueSnapshotBuilder.Build(
                                selection,
                                token),
                            token);
                        token.ThrowIfCancellationRequested();
                        AddPreparedSelection(selection, snapshot);
                        addedCount++;
                    }
                    catch (OperationCanceledException)
                    {
                        throw;
                    }
                    catch (Exception ex)
                    {
                        if (firstFailure == null)
                        {
                            firstFailure = ex;
                        }
                        DiagnosticsLogger.LogException(
                            LogCategories.FileLink,
                            "Queue selection could not be read.",
                            ex);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                scanCancelled = true;
                DiagnosticsLogger.Log(
                    LogCategories.FileLink,
                    "Queue selection scan cancelled.");
            }
            finally
            {
                if (_cancellationSource != null)
                {
                    _cancellationSource.Dispose();
                    _cancellationSource = null;
                }
                _queueScanInProgress = false;
                SetProgressPanelVisible(false);
                _progressBar.Style = ProgressBarStyle.Blocks;
                UpdateNavigationState();
                UpdateUploadButtonState();
                CloseAfterCancellation();
            }

            if (!scanCancelled)
            {
                CompleteSelectionAdd(
                    requestedCount,
                    addedCount,
                    firstFailure);
            }
        }

        internal bool TryAdoptComposeSelections(
            IReadOnlyList<FileLinkSelection> selections)
        {
            if (IsDisposed || Disposing || _closeRequested || _shareFinalized
                || selections == null || selections.Count == 0)
            {
                return false;
            }
            if (InvokeRequired)
            {
                throw new InvalidOperationException("Compose queue adoption requires the wizard thread.");
            }
            var snapshots = new Dictionary<FileLinkSelection, FileLinkQueueNode>();
            var adopted = new List<FileLinkSelection>();
            var existingSelections = new HashSet<FileLinkSelection>(
                _items,
                FileLinkSelection.IdentityComparer);
            try
            {
                foreach (FileLinkSelection selection in selections)
                {
                    if (selection == null
                        || selection.Source != FileLinkSelectionSource.Local
                        || selection.SelectionType != FileLinkSelectionType.File
                        || !existingSelections.Add(selection))
                    {
                        return false;
                    }
                    snapshots.Add(selection, FileLinkQueueSnapshotBuilder.Build(
                        selection, CancellationToken.None));
                }
                foreach (KeyValuePair<FileLinkSelection, FileLinkQueueNode> entry in snapshots)
                {
                    adopted.Add(entry.Key);
                    AddPreparedSelection(entry.Key, entry.Value);
                }
                CompleteSelectionAdd(selections.Count, snapshots.Count, null);
                return true;
            }
            catch (Exception ex)
            {
                RollbackComposeSelections(adopted);
                DiagnosticsLogger.LogException(LogCategories.FileLink,
                    "Compose attachment queue adoption failed.", ex);
                return false;
            }
        }

        internal void RollbackComposeSelections(
            IReadOnlyList<FileLinkSelection> selections)
        {
            foreach (FileLinkSelection selection in selections)
            {
                if (selection == null || _uploadedSelections.Contains(selection))
                {
                    continue;
                }
                _items.Remove(selection);
                _queueSnapshots.Remove(selection);
                SelectionUploadState state;
                if (_selectionStates.TryGetValue(selection, out state))
                {
                    DisposeStateProgressBar(state);
                    _selectionStates.Remove(selection);
                }
            }
            _uploadCompleted = _uploadContext != null
                && _items.All(selection => _uploadedSelections.Contains(selection));
            RebuildQueueView();
            UpdateNavigationState();
            UpdateUploadButtonState();
        }

        private void CompleteSelectionAdd(
            int requestedCount,
            int addedCount,
            Exception firstFailure)
        {
            if (addedCount > 0)
            {
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
                RebuildQueueView();
                _uploadCompleted = false;
                UpdateNavigationState();
                UpdateUploadButtonState();
                if (_uploadContext != null && !_uploadInProgress && !_finalizeInProgress && !_queueScanInProgress
                    && IsHandleCreated && !_closeRequested)
                {
                    BeginInvoke(new Action(async () =>
                    {
                        if (!_uploadCompleted && !_uploadInProgress && !_finalizeInProgress && !_queueScanInProgress
                            && !_closeRequested && !IsDisposed)
                        {
                            await StartUploadAsync();
                        }
                    }));
                }
            }

            if (firstFailure != null)
            {
                MessageBox.Show(
                    string.Format(
                        CultureInfo.CurrentCulture,
                        Strings.FileLinkQueueReadFailedFormat,
                        firstFailure.Message),
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);
            }
        }

    }
}
