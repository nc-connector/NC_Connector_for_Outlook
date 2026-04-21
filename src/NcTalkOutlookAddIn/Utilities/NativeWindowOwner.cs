// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Windows.Forms;

namespace NcTalkOutlookAddIn.Utilities
{
        // Lightweight IWin32Window wrapper for dialog ownership via native HWND.
    internal sealed class NativeWindowOwner : IWin32Window
    {
        private readonly IntPtr _handle;

        internal NativeWindowOwner(IntPtr handle)
        {
            _handle = handle;
        }

        public IntPtr Handle
        {
            get { return _handle; }
        }
    }
}
