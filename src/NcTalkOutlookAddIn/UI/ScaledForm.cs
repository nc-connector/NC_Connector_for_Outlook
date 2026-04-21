// Copyright (c) 2026 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System.Windows.Forms;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    internal abstract class ScaledForm : Form
    {
        protected int ScaleLogical(int value)
        {
            return DpiScaling.ScaleLogical(this, value);
        }
    }
}
