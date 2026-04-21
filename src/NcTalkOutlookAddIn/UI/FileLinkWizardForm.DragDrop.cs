/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    /**
     * Drag/drop and queue-append behavior for the file selection step.
     */
    internal sealed partial class FileLinkWizardForm
    {
        private void AttachFileQueueDropTarget(Control control)
        {            if (control == null)
            {
                return;
            }

            control.AllowDrop = true;
            control.DragEnter += HandleFileListViewDragEnter;
            control.DragOver += HandleFileListViewDragOver;
            control.DragDrop += HandleFileListViewDragDrop;
        }

        private void HandleFileListViewDragEnter(object sender, DragEventArgs e)
        {            if (e == null)
            {
                return;
            }

            e.Effect = ResolveFileDropEffect(e);
        }

        private void HandleFileListViewDragOver(object sender, DragEventArgs e)
        {            if (e == null)
            {
                return;
            }

            e.Effect = ResolveFileDropEffect(e);
        }

        private void HandleFileListViewDragDrop(object sender, DragEventArgs e)
        {            if (e == null)
            {
                return;
            }
            var selections = BuildSelectionsFromFileDropData(e.Data);
            if (selections.Count == 0)
            {
                return;
            }

            AddSelections(selections);
        }

        private static DragDropEffects ResolveFileDropEffect(DragEventArgs e)
        {            if (e == null || e.Data == null)
            {
                return DragDropEffects.None;
            }
            return e.Data.GetDataPresent(DataFormats.FileDrop)
                ? DragDropEffects.Copy
                : DragDropEffects.None;
        }

        private static List<FileLinkSelection> BuildSelectionsFromFileDropData(IDataObject dataObject)
        {
            var selections = new List<FileLinkSelection>();            if (dataObject == null || !dataObject.GetDataPresent(DataFormats.FileDrop))
            {
                return selections;
            }
            var paths = dataObject.GetData(DataFormats.FileDrop) as string[];            if (paths == null || paths.Length == 0)
            {
                return selections;
            }
            for (int i = 0; i < paths.Length; i++)
            {
                string path = paths[i];
                FileLinkSelection selection;
                if (TryBuildSelectionFromPath(path, out selection))
                {
                    selections.Add(selection);
                }
            }
            return selections;
        }

        private bool TryAddSelection(FileLinkSelection selection, HashSet<string> existingPaths)
        {            if (selection == null || string.IsNullOrWhiteSpace(selection.LocalPath))
            {
                return false;
            }
            if (!_attachmentMode && existingPaths != null)
            {
                if (existingPaths.Contains(selection.LocalPath))
                {
                    return false;
                }

                existingPaths.Add(selection.LocalPath);
            }

            _items.Add(selection);

            var listViewItem = new ListViewItem(selection.LocalPath)
            {
                Tag = selection
            };
            listViewItem.UseItemStyleForSubItems = false;
            listViewItem.SubItems.Add(selection.SelectionType == FileLinkSelectionType.File ? Strings.FileLinkWizardTypeFile : Strings.FileLinkWizardTypeFolder);
            listViewItem.SubItems.Add(string.Empty);
            _fileListView.Items.Add(listViewItem);

            var state = new SelectionUploadState(listViewItem);
            _selectionStates[selection] = state;
            return true;
        }

        private static bool SelectionPathExists(FileLinkSelection selection)
        {            if (selection == null || string.IsNullOrWhiteSpace(selection.LocalPath))
            {
                return false;
            }
            return selection.SelectionType == FileLinkSelectionType.Directory
                ? Directory.Exists(selection.LocalPath)
                : File.Exists(selection.LocalPath);
        }

        private static bool TryBuildSelectionFromPath(string path, out FileLinkSelection selection)
        {
            selection = null;
            if (string.IsNullOrWhiteSpace(path))
            {
                return false;
            }
            if (File.Exists(path))
            {
                selection = new FileLinkSelection(FileLinkSelectionType.File, path);
                return true;
            }
            if (Directory.Exists(path))
            {
                selection = new FileLinkSelection(FileLinkSelectionType.Directory, path);
                return true;
            }
            return false;
        }
    }
}

