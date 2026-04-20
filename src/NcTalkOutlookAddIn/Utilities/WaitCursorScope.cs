/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Windows.Forms;

namespace NcTalkOutlookAddIn.Utilities
{
    /**
     * Disposable cursor helper for short-running UI operations.
     */
    internal sealed class WaitCursorScope : IDisposable
    {
        private readonly Cursor _previous;

        internal WaitCursorScope()
        {
            _previous = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;
        }

        public void Dispose()
        {
            Cursor.Current = _previous;
        }
    }
}
