// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Services;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
        // Drag/drop and queue-append behavior for the file selection step.
    internal sealed partial class FileLinkWizardForm
    {
        private void AttachFileQueueDropTarget(Control control)
        {
            if (control == null)
            {
                return;
            }

            control.AllowDrop = true;
            control.DragEnter += HandleFileListViewDragEffect;
            control.DragOver += HandleFileListViewDragEffect;
            control.DragDrop += HandleFileListViewDragDrop;
        }

        private void HandleFileListViewDragEffect(object sender, DragEventArgs e)
        {
            if (e == null)
            {
                return;
            }

            e.Effect = IsWizardBusy
                ? DragDropEffects.None
                : ResolveFileDropEffect(e);
        }

        // WinForms drag-and-drop handlers must stay async void; keep the awaited flow inside this method-level try/catch.
        private async void HandleFileListViewDragDrop(
            object sender,
            DragEventArgs e)
        {
            try
            {
                if (e == null || IsWizardBusy)
                {
                    return;
                }
                var selections = BuildSelectionsFromFileDropData(
                    e.Data);
                if (selections.Count == 0)
                {
                    return;
                }

                await AddSelectionsAsync(selections);
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(
                    LogCategories.FileLink,
                    "Queue drag-and-drop handler failed.",
                    ex);
            }
        }

        private static DragDropEffects ResolveFileDropEffect(DragEventArgs e)
        {
            if (e == null || e.Data == null)
            {
                return DragDropEffects.None;
            }
            return e.Data.GetDataPresent(DataFormats.FileDrop)
                ? DragDropEffects.Copy
                : DragDropEffects.None;
        }

        private static List<FileLinkSelection> BuildSelectionsFromFileDropData(IDataObject dataObject)
        {
            var selections = new List<FileLinkSelection>();
            if (dataObject == null || !dataObject.GetDataPresent(DataFormats.FileDrop))
            {
                return selections;
            }
            var paths = dataObject.GetData(DataFormats.FileDrop) as string[];
            if (paths == null || paths.Length == 0)
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

        private bool TryReserveSelection(
            FileLinkSelection selection,
            HashSet<FileLinkSelection> existingSelections)
        {
            if (selection == null
                || (selection.Source == FileLinkSelectionSource.Local
                    && string.IsNullOrWhiteSpace(selection.LocalPath)))
            {
                return false;
            }
            if (!_attachmentMode && existingSelections != null)
            {
                return existingSelections.Add(selection);
            }
            return true;
        }

        private void AddPreparedSelection(
            FileLinkSelection selection,
            FileLinkQueueNode snapshot)
        {
            if (selection == null || snapshot == null)
            {
                throw new IOException(
                    Strings.FileLinkUploadSourceChanged);
            }
            _items.Add(selection);
            _queueSnapshots.Add(selection, snapshot);
            if (snapshot.IsDirectory && snapshot.Children.Count > 0)
            {
                _expandedQueueFolders.Add(
                    BuildQueueNodeKey(selection, snapshot));
            }

            var state = new SelectionUploadState(selection);
            _selectionStates[selection] = state;
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

