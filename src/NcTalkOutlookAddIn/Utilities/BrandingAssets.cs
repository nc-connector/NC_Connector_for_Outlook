/**
 * Copyright (c) 2025 Bastian Kleinschmidt
 * Licensed under the GNU Affero General Public License v3.0.
 * See LICENSE.txt for details.
 */

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace NcTalkOutlookAddIn.Utilities
{
    /**
     * Central access to embedded branding assets (header banner, app icon).
     */
    internal static class BrandingAssets
    {
        internal const string BrandBlueHex = "#0082C9";
        internal static readonly Color BrandBlue = Color.FromArgb(0, 130, 201); // #0082C9

        private const string HeaderBannerResourceName = "NcTalkOutlookAddIn.Resources.header-solid-blue-164x48.png";
        private const string AppIconResourceName = "NcTalkOutlookAddIn.Resources.app.png";

        private static readonly Lazy<Image> HeaderBanner = new Lazy<Image>(() => LoadEmbeddedBitmap(HeaderBannerResourceName));
        private static readonly Lazy<Image> AppIconImage = new Lazy<Image>(() => LoadEmbeddedBitmap(AppIconResourceName));
        private static readonly Dictionary<int, Icon> IconCache = new Dictionary<int, Icon>();
        private static readonly object IconLock = new object();

        internal static Image HeaderBannerImage
        {
            get { return HeaderBanner.Value; }
        }

        internal static Image AppIconPng
        {
            get { return AppIconImage.Value; }
        }

        internal static Icon GetAppIcon(int size)
        {
            if (size <= 0)
            {
                return null;
            }

            lock (IconLock)
            {
                Icon cached;
                if (IconCache.TryGetValue(size, out cached))
                {
                    return cached;
                }

                Icon icon = CreateIconFromPng(AppIconPng, size);
                if (icon != null)
                {
                    IconCache[size] = icon;
                }
                return icon;
            }
        }

        private static Image LoadEmbeddedBitmap(string resourceName)
        {
            try
            {
                var assembly = Assembly.GetExecutingAssembly();
                using (Stream stream = assembly.GetManifestResourceStream(resourceName))
                {
                    if (stream == null)
                    {
                        return null;
                    }

                    using (var image = Image.FromStream(stream))
                    {
                        return new Bitmap(image);
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "BrandingAssets failed to load embedded bitmap '" + (resourceName ?? string.Empty) + "'.", ex);
                return null;
            }
        }

        private static Icon CreateIconFromPng(Image image, int size)
        {
            if (image == null)
            {
                return null;
            }

            IntPtr hIcon = IntPtr.Zero;
            try
            {
                using (var bitmap = new Bitmap(image, new Size(size, size)))
                {
                    hIcon = bitmap.GetHicon();
                    using (Icon icon = Icon.FromHandle(hIcon))
                    {
                        return (Icon)icon.Clone();
                    }
                }
            }
            catch (Exception ex)
            {
                DiagnosticsLogger.LogException(LogCategories.Core, "BrandingAssets failed to create icon from PNG (size=" + size.ToString() + ").", ex);
                return null;
            }
            finally
            {
                if (hIcon != IntPtr.Zero)
                {
                    DestroyIcon(hIcon);
                }
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern bool DestroyIcon(IntPtr hIcon);
    }
}
