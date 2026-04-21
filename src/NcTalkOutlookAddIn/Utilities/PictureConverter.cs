// Copyright (c) 2025 Bastian Kleinschmidt
// Licensed under the GNU Affero General Public License v3.0.
// See LICENSE.txt for details.

using System.Drawing;
using System.Windows.Forms;

namespace NcTalkOutlookAddIn.Utilities
{
    internal static class PictureConverter
    {
        private sealed class AxHostPictureConverter : AxHost
        {
            internal AxHostPictureConverter() : base(string.Empty)
            {
            }

            internal static stdole.IPictureDisp ImageToPictureDisp(Image image)
            {
                return (stdole.IPictureDisp)GetIPictureDispFromPicture(image);
            }
        }

        internal static stdole.IPictureDisp ToPictureDisp(Image image)
        {
            return AxHostPictureConverter.ImageToPictureDisp(image);
        }
    }
}
