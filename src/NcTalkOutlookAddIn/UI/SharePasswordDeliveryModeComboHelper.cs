// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System.Windows.Forms;
using NcTalkOutlookAddIn.Models;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    internal static class SharePasswordDeliveryModeComboHelper
    {
        internal static void Populate(ComboBox combo)
        {
            if (combo == null)
            {
                return;
            }

            combo.Items.Clear();
            combo.Items.Add(Strings.SharingPasswordDeliveryPlain);
            combo.Items.Add(Strings.SharingPasswordDeliverySecrets);
            combo.SelectedIndex = 0;
        }

        internal static void Select(ComboBox combo, SharePasswordDeliveryMode mode)
        {
            if (combo == null)
            {
                return;
            }

            if (combo.Items.Count == 0)
            {
                Populate(combo);
            }
            combo.SelectedIndex = mode == SharePasswordDeliveryMode.Secrets && combo.Items.Count > 1 ? 1 : 0;
        }

        internal static SharePasswordDeliveryMode GetSelected(ComboBox combo)
        {
            if (combo == null)
            {
                return SharePasswordDeliveryMode.Plain;
            }

            return combo.SelectedIndex == 1 ? SharePasswordDeliveryMode.Secrets : SharePasswordDeliveryMode.Plain;
        }
    }
}
