/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Drawing;
using System.Windows.Forms;
using NcTalkOutlookAddIn.Utilities;

namespace NcTalkOutlookAddIn.UI
{
    /**
     * Lightweight header control that draws the centered branding logo on a solid blue strip.
     */
    internal sealed class BrandedHeader : Control
    {
        private readonly Image _banner;
        private const int HorizontalPadding = 12;

        internal BrandedHeader()
        {
            _banner = BrandingAssets.HeaderBannerImage;

            SetStyle(ControlStyles.UserPaint, true);
            SetStyle(ControlStyles.AllPaintingInWmPaint, true);
            SetStyle(ControlStyles.OptimizedDoubleBuffer, true);
            SetStyle(ControlStyles.ResizeRedraw, true);

            BackColor = BrandingAssets.BrandBlue;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            e.Graphics.Clear(BackColor);

            if (_banner == null || ClientSize.Width <= 0 || ClientSize.Height <= 0)
            {
                return;
            }

            e.Graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            e.Graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;

            Rectangle destination = CalculateContainRectangle(_banner.Size, ClientSize);
            if (!destination.IsEmpty)
            {
                e.Graphics.DrawImage(_banner, destination);
            }
        }

        private static Rectangle CalculateContainRectangle(Size source, Size canvas)
        {
            if (source.Width <= 0 || source.Height <= 0 || canvas.Width <= 0 || canvas.Height <= 0)
            {
                return Rectangle.Empty;
            }

            int availableWidth = Math.Max(0, canvas.Width - (HorizontalPadding * 2));
            int targetHeight = canvas.Height;

            float scaleByHeight = (float)targetHeight / source.Height;
            float scaleByWidth = availableWidth > 0 ? (float)availableWidth / source.Width : scaleByHeight;
            float scale = Math.Min(scaleByHeight, scaleByWidth);

            int drawWidth = (int)Math.Ceiling(source.Width * scale);
            int drawHeight = (int)Math.Ceiling(source.Height * scale);
            int x = (canvas.Width - drawWidth) / 2;
            int y = (canvas.Height - drawHeight) / 2;

            return new Rectangle(x, y, drawWidth, drawHeight);
        }
    }
}
