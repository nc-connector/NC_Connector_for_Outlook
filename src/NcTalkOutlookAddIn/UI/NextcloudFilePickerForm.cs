// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    internal enum NextcloudFilePickerMode
    {
        Files,
        Folder
    }

    internal sealed class NextcloudFilePickerForm : ScaledForm
    {
        private enum PickerNavigationIcon
        {
            Back,
            Forward,
            Up,
            Refresh
        }

        private readonly FileLinkService _service;
        private readonly NextcloudFilePickerMode _mode;
        private readonly UiThemePalette _palette =
            UiThemeManager.DetectPalette();
        private readonly Button _backButton = new Button();
        private readonly Button _forwardButton = new Button();
        private readonly TextBox _filterTextBox = new TextBox();
        private readonly Button _upButton = new Button();
        private readonly Button _refreshButton = new Button();
        private readonly PictureBox _sourceButton = new PictureBox();
        private readonly Panel _addressBar = new Panel();
        private readonly TableLayoutPanel _addressLayout =
            new TableLayoutPanel();
        private readonly FlowLayoutPanel _breadcrumbPanel =
            new FlowLayoutPanel();
        private readonly ListView _itemsView = new ListView();
        private readonly Label _previewLabel = new Label();
        private readonly PictureBox _previewImage = new PictureBox();
        private readonly Label _storageLabel = new Label();
        private readonly Label _statusLabel = new Label();
        private readonly Button _closeButton = new Button();
        private readonly Button _selectButton = new Button();
        private readonly ToolTip _navigationToolTip = new ToolTip();
        private readonly ImageList _icons;
        private readonly NextcloudPickerNavigation _navigation = new NextcloudPickerNavigation();
        private CancellationTokenSource _loadCancellation;
        private CancellationTokenSource _previewCancellation;
        private int _previewVersion;
        private string _currentPath = string.Empty;
        private NextcloudStorageListing _currentListing;
        private long? _usedBytes;
        private long? _availableBytes;
        private bool _loading;

        internal NextcloudFilePickerForm(
            FileLinkService service,
            NextcloudFilePickerMode mode)
        {
            if (service == null)
            {
                throw new ArgumentNullException("service");
            }

            _service = service;
            _mode = mode;
            _icons = FileLinkIconProvider.CreateImageList(
                ScaleLogical(24),
                ScaleLogical(28));
            SelectedEntries = new ReadOnlyCollection<NextcloudStorageEntry>(
                new List<NextcloudStorageEntry>());

            Text = mode == NextcloudFilePickerMode.Files
                ? Strings.NextcloudPickerFilesTitle
                : Strings.NextcloudPickerFolderTitle;
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = false;
            MaximizeBox = true;
            ShowInTaskbar = false;
            ClientSize = new Size(
                ScaleLogical(860),
                ScaleLogical(640));
            MinimumSize = new Size(
                ScaleLogical(720),
                ScaleLogical(520));
            Icon = BrandingAssets.GetAppIcon(32);

            InitializeLayout();
            UiThemeManager.ApplyToForm(this, _navigationToolTip);
            ApplyNavigationTheme();
            Shown += async (s, e) =>
            {
                ApplyNavigationTheme();
                await NavigateToAsync(string.Empty);
            };
        }

        internal ReadOnlyCollection<NextcloudStorageEntry> SelectedEntries
        {
            get;
            private set;
        }

        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            CancelCurrentLoad();
            CancelCurrentPreview();
            DisposePreviewImage();
            DisposeButtonBackgroundImage(_backButton);
            DisposeButtonBackgroundImage(_forwardButton);
            DisposeButtonBackgroundImage(_upButton);
            DisposeButtonBackgroundImage(_refreshButton);
            DisposePictureImage(_sourceButton);
            _navigationToolTip.Dispose();
            _icons.Dispose();
            base.OnFormClosed(e);
        }

        private void InitializeLayout()
        {
            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                Padding = new Padding(0),
                Margin = new Padding(0)
            };
            root.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, ScaleLogical(44)));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            root.RowStyles.Add(new RowStyle(SizeType.Absolute, ScaleLogical(56)));
            Controls.Add(root);

            var navigationBar = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 6,
                RowCount = 1,
                Padding = new Padding(
                    ScaleLogical(8),
                    ScaleLogical(6),
                    ScaleLogical(12),
                    ScaleLogical(6)),
                Margin = new Padding(0)
            };
            navigationBar.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100f));
            navigationBar.ColumnStyles.Add(
                new ColumnStyle(SizeType.Absolute, ScaleLogical(34)));
            navigationBar.ColumnStyles.Add(
                new ColumnStyle(SizeType.Absolute, ScaleLogical(34)));
            navigationBar.ColumnStyles.Add(
                new ColumnStyle(SizeType.Absolute, ScaleLogical(34)));
            navigationBar.ColumnStyles.Add(
                new ColumnStyle(SizeType.Absolute, ScaleLogical(34)));
            navigationBar.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100f));
            navigationBar.ColumnStyles.Add(
                new ColumnStyle(SizeType.Absolute, ScaleLogical(198)));
            root.Controls.Add(navigationBar, 0, 0);

            ConfigureNavigationButton(
                _backButton,
                PickerNavigationIcon.Back,
                Strings.NextcloudPickerNavigateBack);
            _backButton.Click += async (s, e) =>
                await NavigateHistoryAsync(-1);
            navigationBar.Controls.Add(_backButton, 0, 0);

            ConfigureNavigationButton(
                _forwardButton,
                PickerNavigationIcon.Forward,
                Strings.NextcloudPickerNavigateForward);
            _forwardButton.Click += async (s, e) =>
                await NavigateHistoryAsync(1);
            navigationBar.Controls.Add(_forwardButton, 1, 0);

            ConfigureNavigationButton(
                _upButton,
                PickerNavigationIcon.Up,
                Strings.NextcloudPickerNavigateUp);
            _upButton.Click += async (s, e) => await NavigateUpAsync();
            navigationBar.Controls.Add(_upButton, 2, 0);

            ConfigureNavigationButton(
                _refreshButton,
                PickerNavigationIcon.Refresh,
                Strings.NextcloudPickerRefresh);
            _refreshButton.Click += async (s, e) =>
                await RefreshCurrentFolderAsync();
            navigationBar.Controls.Add(_refreshButton, 3, 0);

            _addressBar.Dock = DockStyle.Fill;
            _addressBar.BorderStyle = BorderStyle.FixedSingle;
            _addressBar.Margin = new Padding(
                ScaleLogical(7),
                0,
                ScaleLogical(8),
                0);
            _addressBar.Padding = new Padding(
                ScaleLogical(4),
                0,
                ScaleLogical(4),
                0);
            navigationBar.Controls.Add(_addressBar, 4, 0);

            _addressLayout.Dock = DockStyle.Fill;
            _addressLayout.ColumnCount = 2;
            _addressLayout.RowCount = 1;
            _addressLayout.Margin = new Padding(0);
            _addressLayout.Padding = new Padding(0);
            _addressLayout.RowStyles.Add(
                new RowStyle(SizeType.Percent, 100f));
            _addressLayout.ColumnStyles.Add(
                new ColumnStyle(SizeType.Absolute, ScaleLogical(28)));
            _addressLayout.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100f));
            _addressBar.Controls.Add(_addressLayout);

            _sourceButton.Dock = DockStyle.Fill;
            _sourceButton.Margin = new Padding(0);
            _sourceButton.SizeMode = PictureBoxSizeMode.CenterImage;
            _sourceButton.Image = new Bitmap(
                _icons.Images[FileLinkIconProvider.NextcloudSourceKey]);
            _sourceButton.AccessibleName =
                Strings.FileLinkQueueSourceNextcloud;
            _navigationToolTip.SetToolTip(
                _sourceButton,
                Strings.FileLinkQueueSourceNextcloud);
            _sourceButton.Click += async (s, e) =>
                await NavigateToAsync(string.Empty);
            _addressLayout.Controls.Add(_sourceButton, 0, 0);

            _breadcrumbPanel.Dock = DockStyle.Fill;
            _breadcrumbPanel.AutoScroll = true;
            _breadcrumbPanel.WrapContents = false;
            _breadcrumbPanel.FlowDirection = FlowDirection.LeftToRight;
            _breadcrumbPanel.Margin = new Padding(0);
            _breadcrumbPanel.Padding = new Padding(
                0,
                ScaleLogical(7),
                0,
                0);
            _addressLayout.Controls.Add(_breadcrumbPanel, 1, 0);
            UpdateBreadcrumb(string.Empty);

            _filterTextBox.Dock = DockStyle.Fill;
            _filterTextBox.Margin = new Padding(
                0,
                ScaleLogical(5),
                0,
                ScaleLogical(5));
            _filterTextBox.Text =
                Strings.NextcloudPickerFilterPlaceholder;
            _filterTextBox.ForeColor = _palette.MutedText;
            _filterTextBox.GotFocus += HandleFilterGotFocus;
            _filterTextBox.LostFocus += HandleFilterLostFocus;
            _filterTextBox.TextChanged += HandleFilterChanged;
            navigationBar.Controls.Add(_filterTextBox, 5, 0);

            var split = new SplitContainer
            {
                Dock = DockStyle.Fill,
                Size = new Size(
                    ScaleLogical(860),
                    ScaleLogical(480)),
                Orientation = Orientation.Vertical,
                SplitterWidth = ScaleLogical(1),
                SplitterDistance = ScaleLogical(560),
                Panel1MinSize = ScaleLogical(390),
                Panel2MinSize = ScaleLogical(180),
                Margin = new Padding(0)
            };
            root.Controls.Add(split, 0, 1);

            _itemsView.Dock = DockStyle.Fill;
            _itemsView.BorderStyle = BorderStyle.FixedSingle;
            _itemsView.View = View.Details;
            _itemsView.FullRowSelect = true;
            _itemsView.HideSelection = false;
            _itemsView.MultiSelect = _mode == NextcloudFilePickerMode.Files;
            _itemsView.SmallImageList = _icons;
            _itemsView.Columns.Add(
                Strings.NextcloudPickerColumnName,
                ScaleLogical(300));
            _itemsView.Columns.Add(
                Strings.NextcloudPickerColumnSize,
                ScaleLogical(95));
            _itemsView.Columns.Add(
                Strings.NextcloudPickerColumnModified,
                ScaleLogical(130));
            _itemsView.SelectedIndexChanged += async (s, e) =>
            {
                UpdateSelectButton();
                await UpdatePreviewAsync();
            };
            _itemsView.ItemActivate += async (s, e) =>
                await ActivateSelectedItemAsync();
            _itemsView.KeyDown += async (s, e) =>
                await HandleItemsKeyDownAsync(e);
            _itemsView.Resize += (s, e) => UpdateColumnWidths();
            split.Panel1.Controls.Add(_itemsView);

            _previewLabel.Dock = DockStyle.Fill;
            _previewLabel.Text = Strings.NextcloudPickerNoSelection;
            _previewLabel.TextAlign = ContentAlignment.MiddleCenter;
            _previewLabel.ForeColor = _palette.MutedText;
            _previewLabel.Padding = new Padding(ScaleLogical(16));
            split.Panel2.Controls.Add(_previewLabel);

            _previewImage.Dock = DockStyle.Fill;
            _previewImage.SizeMode = PictureBoxSizeMode.Zoom;
            _previewImage.Padding = new Padding(ScaleLogical(16));
            _previewImage.Visible = false;
            split.Panel2.Controls.Add(_previewImage);

            var footer = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 4,
                RowCount = 1,
                Padding = new Padding(
                    ScaleLogical(12),
                    ScaleLogical(10),
                    ScaleLogical(12),
                    ScaleLogical(10)),
                Margin = new Padding(0)
            };
            footer.ColumnStyles.Add(
                new ColumnStyle(SizeType.Percent, 100f));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            root.Controls.Add(footer, 0, 2);

            _storageLabel.AutoSize = true;
            _storageLabel.Anchor = AnchorStyles.Left;
            footer.Controls.Add(_storageLabel, 0, 0);

            _statusLabel.AutoSize = true;
            _statusLabel.Anchor = AnchorStyles.Right;
            _statusLabel.ForeColor = _palette.MutedText;
            _statusLabel.Margin = new Padding(0, 0, ScaleLogical(12), 0);
            footer.Controls.Add(_statusLabel, 1, 0);

            _closeButton.Text = Strings.NextcloudPickerCloseButton;
            _closeButton.DialogResult = DialogResult.Cancel;
            _closeButton.AutoSize = true;
            _closeButton.Margin = new Padding(0, 0, ScaleLogical(8), 0);
            footer.Controls.Add(_closeButton, 2, 0);

            _selectButton.Text = _mode == NextcloudFilePickerMode.Files
                ? Strings.NextcloudPickerSelectFilesButton
                : Strings.NextcloudPickerSelectFolderButton;
            _selectButton.AutoSize = true;
            _selectButton.Enabled = false;
            _selectButton.Click += async (s, e) => await ConfirmAsync();
            footer.Controls.Add(_selectButton, 3, 0);

            CancelButton = _closeButton;
            AcceptButton = _selectButton;
            UpdateColumnWidths();
        }

        private void ConfigureNavigationButton(
            Button button,
            PickerNavigationIcon icon,
            string accessibleName)
        {
            button.Dock = DockStyle.Fill;
            button.Margin = new Padding(1, 0, 1, 0);
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.AccessibleName = accessibleName;
            button.Enabled = false;
            button.BackgroundImage = CreateNavigationIcon(icon);
            button.BackgroundImageLayout = ImageLayout.Center;
            _navigationToolTip.SetToolTip(button, accessibleName);
        }

        private void ApplyNavigationTheme()
        {
            _filterTextBox.BackColor = _palette.InputBackground;
            _filterTextBox.ForeColor = _palette.MutedText;
            _addressBar.BackColor = _palette.InputBackground;
            _addressLayout.BackColor = _palette.InputBackground;
            _breadcrumbPanel.BackColor = _palette.InputBackground;
            _sourceButton.BackColor = _palette.InputBackground;
            foreach (Button button in new[]
            {
                _backButton,
                _forwardButton,
                _upButton,
                _refreshButton
            })
            {
                button.BackColor = _palette.WindowBackground;
                button.FlatAppearance.BorderSize = 0;
                button.FlatAppearance.MouseOverBackColor =
                    _palette.SelectionBackground;
                button.FlatAppearance.MouseDownBackColor =
                    _palette.SelectionBackground;
            }
        }

        private void DrawNavigationIcon(
            Graphics graphics,
            Rectangle bounds,
            PickerNavigationIcon icon,
            Color color)
        {
            int size = ScaleLogical(16);
            float scale = size / 16f;
            float left = Math.Max(0f, (bounds.Width - size) / 2f);
            float top = Math.Max(0f, (bounds.Height - size) / 2f);
            using (var pen = new Pen(
                color,
                Math.Max(1.4f, 1.7f * scale)))
            {
                graphics.SmoothingMode = SmoothingMode.AntiAlias;
                pen.StartCap = LineCap.Round;
                pen.EndCap = LineCap.Round;
                pen.LineJoin = LineJoin.Round;
                switch (icon)
                {
                    case PickerNavigationIcon.Back:
                        graphics.DrawLine(
                            pen,
                            left + (3f * scale),
                            top + (8f * scale),
                            left + (13f * scale),
                            top + (8f * scale));
                        graphics.DrawLines(
                            pen,
                            new[]
                            {
                                new PointF(left + (7f * scale), top + (4f * scale)),
                                new PointF(left + (3f * scale), top + (8f * scale)),
                                new PointF(left + (7f * scale), top + (12f * scale))
                            });
                        break;
                    case PickerNavigationIcon.Forward:
                        graphics.DrawLine(
                            pen,
                            left + (3f * scale),
                            top + (8f * scale),
                            left + (13f * scale),
                            top + (8f * scale));
                        graphics.DrawLines(
                            pen,
                            new[]
                            {
                                new PointF(left + (9f * scale), top + (4f * scale)),
                                new PointF(left + (13f * scale), top + (8f * scale)),
                                new PointF(left + (9f * scale), top + (12f * scale))
                            });
                        break;
                    case PickerNavigationIcon.Up:
                        graphics.DrawLine(
                            pen,
                            left + (8f * scale),
                            top + (3f * scale),
                            left + (8f * scale),
                            top + (13f * scale));
                        graphics.DrawLines(
                            pen,
                            new[]
                            {
                                new PointF(left + (4f * scale), top + (7f * scale)),
                                new PointF(left + (8f * scale), top + (3f * scale)),
                                new PointF(left + (12f * scale), top + (7f * scale))
                            });
                        break;
                    case PickerNavigationIcon.Refresh:
                        graphics.DrawArc(
                            pen,
                            left + (3f * scale),
                            top + (3f * scale),
                            10f * scale,
                            10f * scale,
                            -65f,
                            285f);
                        using (var brush = new SolidBrush(color))
                        {
                            graphics.FillPolygon(
                                brush,
                                new[]
                                {
                                    new PointF(left + (12.5f * scale), top + (2f * scale)),
                                    new PointF(left + (13.5f * scale), top + (6f * scale)),
                                    new PointF(left + (9.5f * scale), top + (4.5f * scale))
                                });
                        }
                        break;
                }
            }
        }

        private Bitmap CreateNavigationIcon(PickerNavigationIcon icon)
        {
            int size = ScaleLogical(16);
            var bitmap = new Bitmap(size, size);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                DrawNavigationIcon(
                    graphics,
                    new Rectangle(0, 0, size, size),
                    icon,
                    _palette.Text);
            }
            return bitmap;
        }

        private static void DisposeButtonBackgroundImage(Button button)
        {
            if (button == null || button.BackgroundImage == null)
            {
                return;
            }
            button.BackgroundImage.Dispose();
            button.BackgroundImage = null;
        }

        private static void DisposePictureImage(PictureBox picture)
        {
            if (picture == null || picture.Image == null)
            {
                return;
            }
            picture.Image.Dispose();
            picture.Image = null;
        }

        private async Task<bool> LoadFolderAsync(string relativePath)
        {
            bool loaded = false;
            CancelCurrentPreview();
            ShowPreviewText(Strings.NextcloudPickerNoSelection);
            CancelCurrentLoad();
            _loadCancellation = new CancellationTokenSource();
            CancellationToken token = _loadCancellation.Token;
            SetLoading(true, Strings.NextcloudPickerLoading);
            try
            {
                NextcloudStorageListing listing = await Task.Run(
                    () => _service.ListNextcloudDirectory(
                        relativePath,
                        token),
                    token);
                token.ThrowIfCancellationRequested();
                _currentPath = listing.RelativePath;
                _currentListing = listing;
                if (listing.UsedBytes.HasValue)
                {
                    _usedBytes = listing.UsedBytes;
                }
                if (listing.AvailableBytes.HasValue)
                {
                    _availableBytes = listing.AvailableBytes;
                }
                UpdateBreadcrumb(_currentPath);
                PopulateItems();
                UpdateStorageLabel();
                loaded = true;
            }
            catch (OperationCanceledException)
            {
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Nextcloud picker folder load failed.",
                    ex);
                MessageBox.Show(
                    PolicyUiHelper.GetConnectionFailureMessage(ex, ex.Message),
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                if (!token.IsCancellationRequested)
                {
                    SetLoading(false, string.Empty);
                }
            }
            return loaded;
        }

        private async Task NavigateToAsync(string relativePath)
        {
            if (_loading)
            {
                return;
            }
            string targetPath = NextcloudPath.Normalize(relativePath);
            if (_currentListing != null
                && string.Equals(
                    targetPath,
                    _currentPath,
                    StringComparison.Ordinal))
            {
                return;
            }
            if (await LoadFolderAsync(targetPath))
            {
                _navigation.Record(_currentPath);
                UpdateNavigationButtons();
            }
        }

        private async Task NavigateHistoryAsync(int offset)
        {
            if (_loading)
            {
                return;
            }
            int targetIndex;
            string targetPath;
            if (!_navigation.TryGetTarget(offset, out targetIndex, out targetPath))
            {
                return;
            }
            if (await LoadFolderAsync(targetPath))
            {
                _navigation.CompleteHistoryNavigation(targetIndex, _currentPath);
                UpdateNavigationButtons();
            }
        }

        private async Task RefreshCurrentFolderAsync()
        {
            if (_loading || _currentListing == null)
            {
                return;
            }
            await LoadFolderAsync(_currentPath);
            UpdateNavigationButtons();
        }

        private void UpdateNavigationButtons()
        {
            bool enabled = !_loading;
            _backButton.Enabled =
                enabled && _navigation.CanGoBack;
            _forwardButton.Enabled =
                enabled && _navigation.CanGoForward;
            _upButton.Enabled =
                enabled && _currentPath.Length > 0;
            _refreshButton.Enabled =
                enabled && _currentListing != null;
            _sourceButton.Enabled = true;
        }

        private void PopulateItems()
        {
            string filter = GetActiveFilter();
            IEnumerable<NextcloudStorageEntry> entries =
                _currentListing != null
                    ? _currentListing.Entries
                    : Enumerable.Empty<NextcloudStorageEntry>();
            if (!string.IsNullOrWhiteSpace(filter))
            {
                entries = entries.Where(
                    entry => entry.DisplayName.IndexOf(
                        filter,
                        StringComparison.CurrentCultureIgnoreCase) >= 0);
            }

            _itemsView.BeginUpdate();
            try
            {
                _itemsView.Items.Clear();
                foreach (NextcloudStorageEntry entry in entries)
                {
                    string imageKey = entry.IsDirectory
                        ? FileLinkIconProvider.FolderKey
                        : FileLinkIconProvider.GetFileKey(
                            entry.DisplayName);
                    var item = new ListViewItem(
                        entry.DisplayName,
                        imageKey)
                    {
                        Tag = entry
                    };
                    item.SubItems.Add(
                        entry.IsDirectory
                            ? string.Empty
                            : SizeFormatting.FormatBytes(entry.Length));
                    item.SubItems.Add(
                        entry.LastModifiedUtc.HasValue
                            ? entry.LastModifiedUtc.Value
                                .ToLocalTime()
                                .ToString(
                                    "d",
                                    CultureInfo.CurrentCulture)
                            : string.Empty);
                    _itemsView.Items.Add(item);
                }
            }
            finally
            {
                _itemsView.EndUpdate();
            }
            CancelCurrentPreview();
            ShowPreviewText(Strings.NextcloudPickerNoSelection);
            UpdateSelectButton();
        }

        private async Task NavigateUpAsync()
        {
            if (_loading || _currentPath.Length == 0)
            {
                return;
            }
            await NavigateToAsync(NextcloudPath.GetParent(_currentPath));
        }

        private void UpdateBreadcrumb(string relativePath)
        {
            string normalizedPath = NextcloudPath.Normalize(relativePath);
            _breadcrumbPanel.SuspendLayout();
            Control lastPart = null;
            try
            {
                while (_breadcrumbPanel.Controls.Count > 0)
                {
                    Control control = _breadcrumbPanel.Controls[0];
                    _breadcrumbPanel.Controls.RemoveAt(0);
                    control.Dispose();
                }

                string currentPath = string.Empty;
                string[] segments = normalizedPath.Split(
                    new[] { '/' },
                    StringSplitOptions.RemoveEmptyEntries);
                for (int index = 0; index < segments.Length; index++)
                {
                    var separator = new Label
                    {
                        AutoSize = true,
                        Text = "›",
                        BackColor = _palette.InputBackground,
                        ForeColor = _palette.MutedText,
                        Margin = new Padding(
                            ScaleLogical(3),
                            ScaleLogical(2),
                            ScaleLogical(3),
                            0)
                    };
                    _breadcrumbPanel.Controls.Add(separator);

                    currentPath = currentPath.Length == 0
                        ? segments[index]
                        : currentPath + "/" + segments[index];
                    lastPart = AddBreadcrumbPart(
                        segments[index],
                        currentPath,
                        index == segments.Length - 1);
                }
            }
            finally
            {
                _breadcrumbPanel.ResumeLayout(true);
            }
            if (lastPart != null)
            {
                _breadcrumbPanel.ScrollControlIntoView(lastPart);
            }
        }

        private Control AddBreadcrumbPart(
            string text,
            string targetPath,
            bool current)
        {
            if (current)
            {
                var label = new Label
                {
                    AutoSize = true,
                    Text = text,
                    BackColor = _palette.InputBackground,
                    ForeColor = _palette.Text,
                    Font = new Font(
                        Font,
                        FontStyle.Bold),
                    Margin = new Padding(
                        0,
                        ScaleLogical(2),
                        0,
                        0)
                };
                _breadcrumbPanel.Controls.Add(label);
                return label;
            }

            var link = new LinkLabel
            {
                AutoSize = true,
                Text = text,
                BackColor = _palette.InputBackground,
                LinkColor = _palette.LinkText,
                ActiveLinkColor = _palette.LinkText,
                VisitedLinkColor = _palette.LinkText,
                LinkBehavior = LinkBehavior.HoverUnderline,
                Margin = new Padding(
                    0,
                    ScaleLogical(2),
                    0,
                    0)
            };
            string navigationTarget = targetPath;
            link.LinkClicked += async (s, e) =>
            {
                if (!_loading)
                {
                    await NavigateToAsync(navigationTarget);
                }
            };
            _breadcrumbPanel.Controls.Add(link);
            return link;
        }

        private async Task ActivateSelectedItemAsync()
        {
            if (_loading || _itemsView.SelectedItems.Count != 1)
            {
                return;
            }
            NextcloudStorageEntry entry =
                _itemsView.SelectedItems[0].Tag
                as NextcloudStorageEntry;
            if (entry == null)
            {
                return;
            }
            if (entry.IsDirectory)
            {
                await NavigateToAsync(entry.RelativePath);
            }
            else if (_mode == NextcloudFilePickerMode.Files)
            {
                await ConfirmAsync();
            }
        }

        private async Task ConfirmAsync()
        {
            if (_loading)
            {
                return;
            }
            if (_mode == NextcloudFilePickerMode.Files)
            {
                List<NextcloudStorageEntry> files = _itemsView
                    .SelectedItems
                    .Cast<ListViewItem>()
                    .Select(item => item.Tag as NextcloudStorageEntry)
                    .Where(entry => entry != null && !entry.IsDirectory)
                    .ToList();
                if (files.Count == 0)
                {
                    return;
                }
                SelectedEntries = new ReadOnlyCollection<NextcloudStorageEntry>(
                    files);
                DialogResult = DialogResult.OK;
                Close();
                return;
            }

            string folderName = _currentPath.Length == 0
                ? Strings.FileLinkSourceMyNextcloud
                : NextcloudPath.GetName(_currentPath);
            var root = new NextcloudStorageEntry(
                _currentPath,
                folderName,
                true,
                0,
                null);
            SetLoading(true, Strings.NextcloudPickerReadingFolder);
            try
            {
                CancellationToken token = _loadCancellation != null
                    ? _loadCancellation.Token
                    : CancellationToken.None;
                IList<NextcloudStorageEntry> snapshot = await Task.Run(
                    () => _service.SnapshotNextcloudFolder(
                        root,
                        null,
                        token),
                    token);
                token.ThrowIfCancellationRequested();
                var result = new List<NextcloudStorageEntry> { root };
                result.AddRange(snapshot);
                SelectedEntries =
                    new ReadOnlyCollection<NextcloudStorageEntry>(result);
                DialogResult = DialogResult.OK;
                Close();
            }
            catch (OperationCanceledException)
            {
                return;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Nextcloud picker folder snapshot failed.",
                    ex);
                MessageBox.Show(
                    PolicyUiHelper.GetConnectionFailureMessage(ex, ex.Message),
                    Strings.DialogTitle,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                if (!IsDisposed && !Disposing)
                {
                    SetLoading(false, string.Empty);
                }
            }
        }

        private async Task HandleItemsKeyDownAsync(
            KeyEventArgs e)
        {
            if (e.Alt && e.KeyCode == Keys.Left)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                await NavigateHistoryAsync(-1);
                return;
            }
            if (e.Alt && e.KeyCode == Keys.Right)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                await NavigateHistoryAsync(1);
                return;
            }
            if (e.KeyCode == Keys.F5)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                await RefreshCurrentFolderAsync();
                return;
            }
            if (e.KeyCode == Keys.Back)
            {
                e.Handled = true;
                e.SuppressKeyPress = true;
                await NavigateUpAsync();
            }
        }

        private void HandleFilterGotFocus(object sender, EventArgs e)
        {
            if (string.Equals(
                _filterTextBox.Text,
                Strings.NextcloudPickerFilterPlaceholder,
                StringComparison.Ordinal))
            {
                _filterTextBox.Text = string.Empty;
                _filterTextBox.ForeColor = _palette.Text;
            }
        }

        private void HandleFilterLostFocus(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(_filterTextBox.Text))
            {
                _filterTextBox.Text =
                    Strings.NextcloudPickerFilterPlaceholder;
                _filterTextBox.ForeColor = _palette.MutedText;
            }
        }

        private void HandleFilterChanged(object sender, EventArgs e)
        {
            if (_currentListing != null)
            {
                PopulateItems();
            }
        }

        private string GetActiveFilter()
        {
            string value = _filterTextBox.Text ?? string.Empty;
            return string.Equals(
                       value,
                       Strings.NextcloudPickerFilterPlaceholder,
                       StringComparison.Ordinal)
                ? string.Empty
                : value.Trim();
        }

        private async Task UpdatePreviewAsync()
        {
            CancelCurrentPreview();
            int selectionCount = _itemsView.SelectedItems.Count;
            if (selectionCount == 0)
            {
                ShowPreviewText(Strings.NextcloudPickerNoSelection);
                return;
            }
            if (selectionCount > 1)
            {
                ShowPreviewText(string.Format(
                    CultureInfo.CurrentCulture,
                    Strings.NextcloudPickerMultipleSelectionFormat,
                    selectionCount));
                return;
            }

            NextcloudStorageEntry entry =
                _itemsView.SelectedItems[0].Tag
                as NextcloudStorageEntry;
            if (entry == null)
            {
                ShowPreviewText(Strings.NextcloudPickerNoSelection);
                return;
            }
            if (entry.IsDirectory)
            {
                ShowPreviewText(entry.DisplayName);
                return;
            }
            var previewRequest = new NextcloudPickerPreview(_service, entry);

            var cancellation = new CancellationTokenSource();
            _previewCancellation = cancellation;
            int previewVersion = _previewVersion;
            CancellationToken token = cancellation.Token;
            ShowPreviewText(Strings.NextcloudPickerPreviewLoading);
            try
            {
                Bitmap preview = await Task.Run(
                    () => previewRequest.Load(token),
                    token);
                if (!IsCurrentPreview(
                    cancellation,
                    previewVersion))
                {
                    if (preview != null)
                    {
                        preview.Dispose();
                    }
                    return;
                }
                if (preview == null)
                {
                    ShowPreviewText(previewRequest.OriginalImageExceedsLimit
                        ? Strings.NextcloudPickerPreviewSkipped
                        : Strings.NextcloudPickerNoPreview);
                    return;
                }
                ShowPreviewImage(preview);
            }
            catch (OperationCanceledException)
            {
                return;
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Nextcloud picker preview failed.",
                    ex);
                if (IsCurrentPreview(
                    cancellation,
                    previewVersion))
                {
                    ShowPreviewText(
                        Strings.NextcloudPickerPreviewLoadFailed);
                }
            }
            finally
            {
                if (ReferenceEquals(
                    _previewCancellation,
                    cancellation))
                {
                    _previewCancellation = null;
                }
                cancellation.Dispose();
            }
        }

        private bool IsCurrentPreview(
            CancellationTokenSource cancellation,
            int previewVersion)
        {
            return !IsDisposed
                   && !Disposing
                   && ReferenceEquals(
                       _previewCancellation,
                       cancellation)
                   && previewVersion == _previewVersion
                   && !cancellation.IsCancellationRequested;
        }

        private void ShowPreviewText(string text)
        {
            DisposePreviewImage();
            _previewImage.Visible = false;
            _previewLabel.Text = text ?? string.Empty;
            _previewLabel.Visible = true;
            _previewLabel.BringToFront();
        }

        private void ShowPreviewImage(Bitmap image)
        {
            if (image == null)
            {
                ShowPreviewText(Strings.NextcloudPickerPreviewLoadFailed);
                return;
            }
            if (IsDisposed || Disposing)
            {
                image.Dispose();
                return;
            }

            DisposePreviewImage();
            _previewImage.Image = image;
            _previewLabel.Visible = false;
            _previewImage.Visible = true;
            _previewImage.BringToFront();
        }

        private void DisposePreviewImage()
        {
            Image image = _previewImage.Image;
            _previewImage.Image = null;
            if (image != null)
            {
                image.Dispose();
            }
        }

        private void UpdateSelectButton()
        {
            if (_loading)
            {
                _selectButton.Enabled = false;
                return;
            }
            _selectButton.Enabled =
                _mode == NextcloudFilePickerMode.Folder
                || _itemsView.SelectedItems
                    .Cast<ListViewItem>()
                    .Select(item => item.Tag as NextcloudStorageEntry)
                    .Any(entry => entry != null && !entry.IsDirectory);
        }

        private void UpdateStorageLabel()
        {
            if (!_usedBytes.HasValue)
            {
                _storageLabel.Text = string.Empty;
                return;
            }
            if (_availableBytes.HasValue)
            {
                long total = checked(
                    _usedBytes.Value + _availableBytes.Value);
                _storageLabel.Text = string.Format(
                    CultureInfo.CurrentCulture,
                    Strings.NextcloudPickerStorageOfFormat,
                    SizeFormatting.FormatBytes(_usedBytes.Value),
                    SizeFormatting.FormatBytes(total));
                return;
            }
            _storageLabel.Text = string.Format(
                CultureInfo.CurrentCulture,
                Strings.NextcloudPickerStorageUsedFormat,
                SizeFormatting.FormatBytes(_usedBytes.Value));
        }

        private void SetLoading(bool loading, string status)
        {
            _loading = loading;
            _statusLabel.Text = status ?? string.Empty;
            _itemsView.Enabled = !loading;
            _filterTextBox.Enabled = !loading;
            UpdateNavigationButtons();
            UseWaitCursor = loading;
            UpdateSelectButton();
        }

        private void CancelCurrentLoad()
        {
            if (_loadCancellation == null)
            {
                return;
            }
            try
            {
                _loadCancellation.Cancel();
            }
            catch (ObjectDisposedException)
            {
            }
            _loadCancellation.Dispose();
            _loadCancellation = null;
        }

        private void CancelCurrentPreview()
        {
            _previewVersion++;
            CancellationTokenSource cancellation =
                _previewCancellation;
            _previewCancellation = null;
            if (cancellation == null)
            {
                return;
            }
            try
            {
                cancellation.Cancel();
            }
            catch (ObjectDisposedException)
            {
            }
        }

        private void UpdateColumnWidths()
        {
            if (_itemsView.Columns.Count < 3)
            {
                return;
            }
            int available = Math.Max(
                ScaleLogical(360),
                _itemsView.ClientSize.Width - ScaleLogical(6));
            int sizeWidth = ScaleLogical(96);
            int modifiedWidth = ScaleLogical(132);
            _itemsView.Columns[0].Width = Math.Max(
                ScaleLogical(160),
                available - sizeWidth - modifiedWidth);
            _itemsView.Columns[1].Width = sizeWidth;
            _itemsView.Columns[2].Width = modifiedWidth;
        }

    }
}
