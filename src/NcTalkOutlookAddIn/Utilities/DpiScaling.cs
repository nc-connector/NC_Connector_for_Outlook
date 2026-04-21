// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System;
using System.Windows.Forms;

namespace NcTalkOutlookAddIn.Utilities
{
        // Shared DPI scaling helper for WinForms dialogs.
    internal static class DpiScaling
    {
        internal static int ScaleLogical(Control control, int value)
        {
            int dpi = control != null && control.DeviceDpi > 0 ? control.DeviceDpi : 96;
            return (int)Math.Round(value * (dpi / 96f));
        }
    }
}
